VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form diaAPe02a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Vendor Credit Or Debit Memo"
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6990
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   6990
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtFrt 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   600
      TabIndex        =   16
      Tag             =   "1"
      ToolTipText     =   "Freight and Misc"
      Top             =   5880
      Width           =   1095
   End
   Begin VB.TextBox txtTax 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   600
      TabIndex        =   17
      Tag             =   "1"
      ToolTipText     =   "Sales Tax For Invoice"
      Top             =   6240
      Width           =   1095
   End
   Begin VB.TextBox txtAcctAmt 
      Height          =   285
      Index           =   4
      Left            =   600
      TabIndex        =   14
      Tag             =   "1"
      Top             =   5520
      Width           =   1095
   End
   Begin VB.TextBox txtAcctAmt 
      Height          =   285
      Index           =   3
      Left            =   600
      TabIndex        =   12
      Tag             =   "1"
      Top             =   5160
      Width           =   1095
   End
   Begin VB.TextBox txtAcctAmt 
      Height          =   285
      Index           =   2
      Left            =   600
      TabIndex        =   10
      Tag             =   "1"
      Top             =   4800
      Width           =   1095
   End
   Begin VB.TextBox txtAcctAmt 
      Height          =   285
      Index           =   1
      Left            =   600
      TabIndex        =   8
      Tag             =   "1"
      Top             =   4440
      Width           =   1095
   End
   Begin VB.TextBox txtAcctAmt 
      Height          =   285
      Index           =   0
      Left            =   600
      TabIndex        =   6
      Tag             =   "1"
      Top             =   4080
      Width           =   1095
   End
   Begin VB.ComboBox cmbAct 
      Height          =   315
      Index           =   4
      Left            =   2040
      Sorted          =   -1  'True
      TabIndex        =   15
      Tag             =   "3"
      Top             =   5520
      Width           =   1695
   End
   Begin VB.ComboBox cmbAct 
      Height          =   315
      Index           =   3
      Left            =   2040
      Sorted          =   -1  'True
      TabIndex        =   13
      Tag             =   "3"
      Top             =   5160
      Width           =   1695
   End
   Begin VB.ComboBox cmbAct 
      Height          =   315
      Index           =   2
      Left            =   2040
      Sorted          =   -1  'True
      TabIndex        =   11
      Tag             =   "3"
      Top             =   4800
      Width           =   1695
   End
   Begin VB.ComboBox cmbAct 
      Height          =   315
      Index           =   1
      Left            =   2040
      Sorted          =   -1  'True
      TabIndex        =   9
      Tag             =   "3"
      Top             =   4440
      Width           =   1695
   End
   Begin VB.ComboBox cmbAct 
      Height          =   315
      Index           =   0
      Left            =   2040
      Sorted          =   -1  'True
      TabIndex        =   7
      Tag             =   "3"
      Top             =   4080
      Width           =   1695
   End
   Begin VB.ComboBox txtPst 
      Height          =   315
      Left            =   4560
      TabIndex        =   4
      Tag             =   "4"
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox txtCmt 
      Height          =   855
      Left            =   1560
      TabIndex        =   18
      Tag             =   "9"
      ToolTipText     =   "Optional Comment"
      Top             =   7320
      Width           =   5295
   End
   Begin VB.CommandButton cmdPst 
      Caption         =   "&Post"
      Enabled         =   0   'False
      Height          =   315
      Left            =   6000
      TabIndex        =   19
      ToolTipText     =   "Save and Post This Item"
      Top             =   480
      Width           =   875
   End
   Begin VB.ComboBox cmbTyp 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   4920
      TabIndex        =   1
      Tag             =   "8"
      ToolTipText     =   "Select CM Or DM From  List"
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ComboBox txtMemoDate 
      Height          =   315
      Left            =   2040
      TabIndex        =   3
      Tag             =   "4"
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox txtAmt 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2160
      TabIndex        =   5
      Tag             =   "1"
      ToolTipText     =   "Negative Number For Credit Memo"
      Top             =   8400
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox cmbVnd 
      Height          =   315
      Left            =   2040
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Select Or Enter A Vendor"
      Top             =   480
      Width           =   1555
   End
   Begin VB.TextBox txtInv 
      Height          =   285
      Left            =   2040
      TabIndex        =   2
      Tag             =   "3"
      ToolTipText     =   "20 Char Invoice Number (Min 4). Must Be Unique To The Vendor"
      Top             =   2280
      Width           =   2775
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   6000
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   21
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
      PictureUp       =   "diaAPe02a.frx":0000
      PictureDn       =   "diaAPe02a.frx":0146
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   4560
      Top             =   0
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   8355
      FormDesignWidth =   6990
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Invoice Amount"
      Height          =   285
      Index           =   11
      Left            =   2040
      TabIndex        =   43
      Top             =   6720
      Width           =   1905
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Freight"
      Height          =   255
      Index           =   9
      Left            =   2040
      TabIndex        =   42
      Top             =   5880
      Width           =   495
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Tax"
      Height          =   255
      Index           =   10
      Left            =   2040
      TabIndex        =   41
      Top             =   6240
      Width           =   735
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   6960
      Y1              =   6600
      Y2              =   6600
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Distributes Invoice Amount"
      Height          =   285
      Index           =   6
      Left            =   240
      TabIndex        =   40
      Top             =   3360
      Width           =   1905
   End
   Begin VB.Label lblTot 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   600
      TabIndex        =   39
      Top             =   6720
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
      Height          =   255
      Index           =   2
      Left            =   720
      TabIndex        =   38
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   3960
      TabIndex        =   37
      Top             =   5520
      Width           =   2775
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   3960
      TabIndex        =   36
      Top             =   5160
      Width           =   2775
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   3960
      TabIndex        =   35
      Top             =   4800
      Width           =   2775
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   3960
      TabIndex        =   34
      Top             =   4440
      Width           =   2775
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Remit To:"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   33
      Top             =   840
      Width           =   1185
   End
   Begin VB.Label lblNme 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   1155
      Left            =   2040
      TabIndex        =   32
      Top             =   840
      Width           =   2775
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   0
      Left            =   3960
      TabIndex        =   31
      Top             =   4080
      Width           =   2775
   End
   Begin VB.Label lblDue 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   4920
      TabIndex        =   30
      Top             =   1680
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Account"
      Height          =   255
      Index           =   8
      Left            =   2400
      TabIndex        =   29
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Posting Date"
      Height          =   285
      Index           =   4
      Left            =   3480
      TabIndex        =   28
      Top             =   2760
      Width           =   945
   End
   Begin VB.Label lblPst 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   4920
      TabIndex        =   27
      Top             =   1320
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks:"
      Height          =   285
      Index           =   13
      Left            =   240
      TabIndex        =   26
      Top             =   7320
      Width           =   1200
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Credit/Debit Memo Date"
      Height          =   405
      Index           =   7
      Left            =   240
      TabIndex        =   25
      Top             =   2760
      Width           =   1665
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice Amount"
      Height          =   285
      Index           =   5
      Left            =   600
      TabIndex        =   24
      Top             =   8400
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Vendor"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   23
      Top             =   480
      Width           =   1185
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Credit/Debit Number"
      Height          =   315
      Index           =   3
      Left            =   240
      TabIndex        =   22
      Top             =   2280
      Width           =   1665
   End
End
Attribute VB_Name = "diaAPe02a"
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
'
' diaAPe02a - Vendor Credit Or Debit Memo.
'
' Notes:
'
' Created: (nth)
' Revisions:
'   12/20/02 (nth) Removed option to choose CM or DM per JLH.
'   12/26/06 (nth) Allow negative invoice amounts (credit memo) per JLH.
'   08/25/04 (nth) Post to correct journal via txtpst.
'   09/29/04 (nth) Added remit to.
'
'*************************************************************************************

Option Explicit

Dim bCancel As Byte
Dim bOnLoad As Byte
Dim bGoodVendor As Byte
Dim sJournalID As String
Dim sDefaultAccount As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub ManageBoxes(bEnabled As Byte)
   On Error Resume Next
   If bEnabled = 1 Then
      txtInv.enabled = True
      txtMemoDate.enabled = True
      txtPst.enabled = True
      txtAmt.enabled = True
      txtCmt.enabled = True
      'cmbAct.enabled = True
   Else
      txtInv.enabled = False
      txtMemoDate.enabled = False
      txtPst.enabled = False
      txtAmt.enabled = False
      txtCmt.enabled = False
      cmdPst.enabled = False
      'cmbAct.enabled = False
   End If
   txtInv = ""
   txtAmt = "0.00"
   txtCmt = ""
End Sub

Private Sub cmbAct_Click(Index As Integer)
    If Trim(txtAcctAmt(Index)) = "" Then
        lblDsc(Index) = ""
        cmbAct(Index) = ""
    Else
        If CCur(txtAcctAmt(Index)) = 0 Then
            lblDsc(Index) = ""
            cmbAct(Index) = ""
        Else
            lblDsc(Index) = UpdateActDesc(cmbAct(Index), lblDsc(Index), True)
        End If
    End If
    
End Sub

Private Sub cmbAct_GotFocus(Index As Integer)
    SelectFormat Me
End Sub

Private Sub cmbAct_LostFocus(Index As Integer)
    If Trim(txtAcctAmt(Index)) = "" Then
        lblDsc(Index) = ""
        cmbAct(Index) = ""
    Else
        If IsNumeric(txtAcctAmt(Index)) = 0 Then
            If CCur(txtAcctAmt(Index)) = 0 Then
                lblDsc(Index) = ""
                cmbAct(Index) = ""
            Else
                lblDsc(Index) = UpdateActDesc(cmbAct(Index), lblDsc(Index), True)
            End If
        End If
    End If
End Sub

'Private Sub cmbAct_Click()
'   FindAccount Me
'End Sub

'Private Sub cmbAct_LostFocus()
'   cmbAct = CheckLen(cmbAct, 12)
'   If Len(Trim(cmbAct)) Then FindAccount Me
'End Sub

Private Sub cmbTyp_Change()
   cmbTyp = CheckLen(cmbTyp, 2)
   If cmbTyp <> "CM" And cmbTyp <> "DM" Then cmbTyp = "CM"
   
End Sub

Private Sub cmbVnd_Click()
   bGoodVendor = FindVendor(Me, , , True)
End Sub

Private Sub cmbVnd_LostFocus()
   cmbVnd = CheckLen(cmbVnd, 10)
   If bCancel Then Exit Sub
   bGoodVendor = FindVendor(Me, , , True)
   If bGoodVendor = 1 Then
      ManageBoxes 1
   Else
      ManageBoxes 0
   End If
End Sub

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, _
                             X As Single, Y As Single)
   bCancel = True
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      SelectHelpTopic Me, Caption
      cmdHlp = False
      MouseCursor 0
   End If
End Sub

Private Sub cmdPst_Click()
   Dim b As Byte
   'If Trim(cmbAct) = "" Or lblDsc.ForeColor = ES_RED Then
   '   MsgBox "Requires A Valid Account.", _
   '      vbInformation, Caption
   '   On Error Resume Next
   '   cmbAct.SetFocus
   '   Exit Sub
   'End If
   b = GetVendorInvoice()
   If b = 0 Then
      PostThisMemo
   Else
      MsgBox "That Invoice Number Is In Use By The Current Vendor.", _
         vbInformation, Caption
   End If
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      sDefaultAccount = ""
      CurrentJournal "PJ", ES_SYSDATE, sJournalID
      FillCombo
      bOnLoad = False
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   sCurrForm = Caption
   bOnLoad = True
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If bGoodVendor Then
      cUR.CurrentVendor = cmbVnd
      SaveCurrentSelections
   End If
   FormUnload
   Set diaAPp02a = Nothing
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtMemoDate = Format(Now, "mm/dd/yy")
   lblPst = Format(Now, "mm/dd/yy")
   lblDue = Format(Now + 30, "mm/dd/yy")
   txtPst = lblPst
   txtAmt = "0.00"
   
End Sub

Private Sub FillCombo()
   Dim b As Byte

   On Error GoTo DiaErr1
   AddComboStr cmbTyp.hWnd, "CM"
   AddComboStr cmbTyp.hWnd, "DM"
   cmbTyp = cmbTyp.List(0)
   FillVendors Me
   cmbVnd = cUR.CurrentVendor
   bGoodVendor = FindVendor(Me)
   FillAccounts
   Exit Sub
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub



Private Sub lblNme_Change()
   If Left(lblNme, 6) = "*** No" Then
      lblNme.ForeColor = ES_RED
      bGoodVendor = 0
   Else
      lblNme.ForeColor = Me.ForeColor
      bGoodVendor = 1
   End If
End Sub



Private Sub txtAcctAmt_GotFocus(Index As Integer)
    SelectFormat Me
End Sub

Private Sub txtAcctAmt_LostFocus(Index As Integer)
   If Trim(txtAcctAmt(Index).Text) = "" Then
      UpdateTotals
      Exit Sub
   End If
   
   If Not IsNumeric(txtAcctAmt(Index).Text) Then
      MsgBox "Amounts must be numeric"
      txtAcctAmt(Index).SetFocus
      Exit Sub
   End If
   If Trim(txtAcctAmt(Index)) <> "" Then
      txtAcctAmt(Index) = Format(txtAcctAmt(Index), CURRENCYMASK)
      If CCur(txtAcctAmt(Index)) = 0 Then
         cmbAct(Index) = ""
         lblDsc(Index) = ""
      Else
        If cmbAct(Index) = "" Then cmbAct(Index) = sDefaultAccount
      End If
   Else
      cmbAct(Index) = ""
      lblDsc(Index) = ""
   End If
   UpdateTotals
End Sub

Private Sub txtFrt_LostFocus()
   If Trim(Trim(txtFrt.Text)) = "" Then
      UpdateTotals
      Exit Sub
   End If
   
   If Not IsNumeric(Trim(txtFrt.Text)) Then
      MsgBox "Amounts must be numeric"
      txtFrt.SetFocus
      Exit Sub
   End If
   
   If Trim(txtFrt) <> "" Then
      txtFrt = Format(Trim(txtFrt), CURRENCYMASK)
   End If
   UpdateTotals

End Sub
Private Sub txtTax_LostFocus()
   If Trim(txtTax.Text) = "" Then
      UpdateTotals
      Exit Sub
   End If
   
   If Not IsNumeric(Trim(txtTax.Text)) Then
      MsgBox "Amounts must be numeric"
      txtFrt.SetFocus
      Exit Sub
   End If
   
   If Trim(txtTax) <> "" Then
      txtTax = Format(Trim(txtTax), CURRENCYMASK)
   End If
   UpdateTotals

End Sub


Private Sub txtAmt_LostFocus()
   If CCur(Trim(txtAmt)) < 0 Then
      cmbTyp = "CM"
   Else
      cmbTyp = "DM"
   End If
   txtAmt = CheckLen(txtAmt, 11)
   txtAmt = Format(CCur(txtAmt), "#####0.00")
   UpdateTotals
End Sub

Private Sub txtCmt_LostFocus()
   txtCmt = CheckLen(txtCmt, 255)
   txtCmt = StrCase(txtCmt, ES_FIRSTWORD)
End Sub

Private Sub txtInv_LostFocus()
   txtInv = CheckLen(txtInv, 20)
   If Len(Trim(txtInv)) > 0 Then
      cmdPst.enabled = True
   End If
End Sub

Private Sub txtPdt_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtMemoDate_LostFocus()
   txtMemoDate = CheckDate(txtMemoDate)
   If Format(txtMemoDate, "yyyy/mm/dd") > Format(lblPst, "yyyy/mm/dd") Then
      Beep
      txtMemoDate = lblPst
   End If
End Sub

Private Function GetVendorInvoice() As Byte
   Dim AdoInv As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT VINO,VIVENDOR FROM VihdTable " _
          & "WHERE VINO='" & txtInv & "' AND " _
          & "VIVENDOR='" & Compress(cmbVnd) & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, AdoInv, ES_FORWARD)
   If bSqlRows Then GetVendorInvoice = 1 Else _
                                       GetVendorInvoice = 0
   'AdoInv.Cancel
   Set AdoInv = Nothing
   Exit Function
DiaErr1:
   sProcName = "getvendorinv"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Function

Private Sub PostThisMemo()
   Dim b As Byte
   Dim bResponse As Byte
   Dim iTrans As Long
   Dim iRef As Integer
   Dim cTotal As Currency
   Dim sMsg As String
   Dim sType As String
   Dim sVendor As String
   Dim sApAcct As String
   Dim sRvAcct As String
   Dim sNewAcct As String
   Dim sToday As String
   Dim sInv As String
   
   Dim sTxAcct As String
   Dim sFrAcct As String
   
   Dim CLineAmt As Currency
   Dim i As Integer
   Dim cFREIGHT As Currency
   Dim cTax As Currency
   
   
   On Error GoTo DiaErr1
   
   
'   If lblTot <> txtAmt Then
'      sMsg = "Distribution Total Does Not Match Check Amount."
'      MsgBox sMsg, vbInformation, Caption
'      txtAcctAmt(0).SetFocus
'      Exit Sub
'   End If
   For b = 0 To 4
      If Trim(txtAcctAmt(b)) <> "" Then
         If Trim(cmbAct(b)) = "" And CCur(txtAcctAmt(b)) > 0 Then
            sMsg = "One Or More Distributions Are Missing An Account."
            MsgBox sMsg, vbInformation, Caption
            cmbAct(b).SetFocus
            Exit Sub
         End If
      End If
   Next
   
   sVendor = Compress(cmbVnd)
   sInv = Trim(txtInv)
   sType = Trim(cmbTyp)
   'sNewAcct = Compress(cmbAct(0))
   sToday = Format(ES_SYSDATE, "mm/dd/yy")
   sJournalID = GetOpenJournal("PJ", txtPst)
   If sJournalID <> "" Then
      iTrans = GetNextTransaction(sJournalID)
   Else
      MsgBox "No Open Journal Purchases Journal Found For " _
         & txtPst & " .", vbInformation, Caption
      txtPst.SetFocus
      Exit Sub
   End If
   b = GetDBAccounts(sApAcct, sRvAcct, sTxAcct, sFrAcct)
   If sApAcct = "" Or sRvAcct = "" Then
      MsgBox "Missing One Or Both Required Accounts. & vbCr" _
         & "Can't Post This Memo.", _
         vbInformation, Caption
      Exit Sub
   End If
   
   If sType = "DM" Then
      sMsg = "Are You Ready To Post The Vendor Debit Memo?.."
   Else
      sMsg = "Are You Ready To Post The Vendor Credit Memo?.."
   End If
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      MouseCursor 13
      iTrans = GetNextTransaction(sJournalID)
      
      cTotal = CCur(Trim(lblTot))
      cmdPst.enabled = False
      
      If (Trim(txtFrt) <> "") Then
         cFREIGHT = CCur(Trim(txtFrt))
      Else
         cFREIGHT = 0
      End If
      
      If (Trim(txtTax) <> "") Then
         cTax = CCur(Trim(txtTax))
      Else
         cTax = 0
      End If
      
      
      On Error Resume Next
      clsADOCon.BeginTrans
      clsADOCon.ADOErrNum = 0
      
      sSql = "INSERT INTO VihdTable (VINO,VIVENDOR,VIDATE," _
             & "VIDUE,VIDTRECD,VIDUEDATE,VIFREIGHT,VITAX,VICOMT,VITYPE) VALUES('" _
             & txtInv & "','" _
             & sVendor & "','" _
             & txtMemoDate & "'," _
             & cTotal & ",'" _
             & lblPst & "','" _
             & lblDue & "'," _
             & cFREIGHT & "," _
             & cTax & ",'" _
             & Trim(txtCmt) & "','" _
             & sType & "')"
      clsADOCon.ExecuteSql sSql
      
'      For I = 0 To 4
'        CLineAmt = CCur(txtAcctAmt(I) & "0")
'        sNewAcct = Compress(cmbAct(I))
'        If CLineAmt <> 0 Then
'            sSql = "INSERT INTO ViitTable (VITNO,VITVENDOR," _
'                 & "VITQTY,VITCOST,VITACCOUNT) VALUES('" _
'                 & txtInv & "','" _
'                 & sVendor & "'," _
'                 & "1," _
'                 & CLineAmt & ",'" _
'                 & sNewAcct & "')"
'            clsADOCon.ExecuteSQL sSql
'        End If
'      Next I
      
      
      If iTrans > 0 Then
         ' for all the
         For i = 0 To 4
            CLineAmt = CCur(txtAcctAmt(i) & "0")
            sNewAcct = Compress(cmbAct(i))
            If CLineAmt <> 0 Then
                sSql = "INSERT INTO ViitTable (VITNO,VITITEM,VITVENDOR," _
                     & "VITQTY,VITCOST,VITACCOUNT) VALUES('" _
                     & txtInv & "','" & CStr(i + 1) & "','" _
                     & sVendor & "'," _
                     & "1," _
                     & CLineAmt & ",'" _
                     & sNewAcct & "')"
                clsADOCon.ExecuteSql sSql
            
               'cTotal = Abs(cTotal)
               If sType = "DM" Then
                  'Credit AP
                  iRef = iRef + 1
                  sSql = "INSERT INTO JritTable (DCHEAD,DCTRAN,DCREF," _
                         & "DCCREDIT,DCACCTNO,DCDATE,DCVENDOR,DCVENDORINV) VALUES('" _
                         & Trim(sJournalID) & "'," _
                         & iTrans & "," _
                         & iRef & "," _
                         & CLineAmt & ",'" _
                         & sApAcct & "','" _
                         & lblPst & "','" _
                         & sVendor & "','" _
                         & Trim(txtInv) & "')"
                  clsADOCon.ExecuteSql sSql
                  'Debit
                  iRef = iRef + 1
                  sSql = "INSERT INTO JritTable (DCHEAD,DCTRAN,DCREF," _
                         & "DCDEBIT,DCACCTNO,DCDATE,DCVENDOR,DCVENDORINV) VALUES('" _
                         & Trim(sJournalID) & "'," _
                         & iTrans & "," _
                         & iRef & "," _
                         & CLineAmt & ",'" _
                         & sNewAcct & "','" _
                         & lblPst & "','" _
                         & sVendor & "','" _
                         & Trim(txtInv) & "')"
                  clsADOCon.ExecuteSql sSql
               Else
                  'Credit AP
                  'Credit memo
                  ' 9/20/2012 Change the negative value to positive
                  
                  iRef = iRef + 1
                  sSql = "INSERT INTO JritTable (DCHEAD,DCTRAN,DCREF," _
                         & "DCDEBIT,DCACCTNO,DCDATE,DCVENDOR,DCVENDORINV) VALUES('" _
                         & Trim(sJournalID) & "'," _
                         & iTrans & "," _
                         & iRef & "," _
                         & Abs(CLineAmt) & ",'" _
                         & sApAcct & "','" _
                         & lblPst & "','" _
                         & sVendor & "','" _
                         & Trim(txtInv) & "')"
                  clsADOCon.ExecuteSql sSql
                  
                  iRef = iRef + 1
                  sSql = "INSERT INTO JritTable (DCHEAD,DCTRAN,DCREF," _
                         & "DCCREDIT,DCACCTNO,DCDATE,DCVENDOR,DCVENDORINV) VALUES('" _
                         & Trim(sJournalID) & "'," _
                         & iTrans & "," _
                         & iRef & "," _
                         & Abs(CLineAmt) & ",'" _
                         & sNewAcct & "','" _
                         & lblPst & "','" _
                         & sVendor & "','" _
                         & Trim(txtInv) & "')"
                  clsADOCon.ExecuteSql sSql
               End If
            End If ' Only if there is amount distributed.
         Next ' For all the distribution combo boxes
      
      
         ' Tax
         If cTax > 0 Then
            ' Credit
            iRef = iRef + 1
            sSql = "INSERT INTO JritTable (DCHEAD,DCTRAN,DCREF," _
                   & "DCCREDIT,DCACCTNO,DCDATE,DCVENDOR,DCVENDORINV) " _
                   & "VALUES('" _
                   & Trim(sJournalID) & "'," _
                   & iTrans & "," _
                   & iRef & "," _
                   & cTax & ",'" _
                   & sApAcct & "','" _
                   & lblPst & "','" _
                   & sVendor & "','" _
                   & Trim(txtInv) & "')"
            clsADOCon.ExecuteSql sSql
            
            ' Debit
            iRef = iRef + 1
            sSql = "INSERT INTO JritTable (DCHEAD,DCTRAN,DCREF," _
                   & "DCDEBIT,DCACCTNO,DCDATE,DCVENDOR,DCVENDORINV) " _
                   & "VALUES('" _
                   & Trim(sJournalID) & "'," _
                   & iTrans & "," _
                   & iRef & "," _
                   & cTax & ",'" _
                   & sTxAcct & "','" _
                   & lblPst & "','" _
                   & sVendor & "','" _
                   & Trim(txtInv) & "')"
            clsADOCon.ExecuteSql sSql
         End If
         
         'Freight
         If cFREIGHT > 0 Then
            ' Credit
            iRef = iRef + 1
            sSql = "INSERT INTO JritTable (DCHEAD,DCTRAN,DCREF," _
                   & "DCCREDIT,DCACCTNO,DCDATE,DCVENDOR,DCVENDORINV) " _
                   & "VALUES('" _
                   & Trim(sJournalID) & "'," _
                   & iTrans & "," _
                   & iRef & "," _
                   & cFREIGHT & ",'" _
                   & sApAcct & "','" _
                   & lblPst & "','" _
                   & sVendor & "','" _
                   & Trim(txtInv) & "')"
            clsADOCon.ExecuteSql sSql
            
            ' Debit
            iRef = iRef + 1
            sSql = "INSERT INTO JritTable (DCHEAD,DCTRAN,DCREF," _
                   & "DCDEBIT,DCACCTNO,DCDATE,DCVENDOR,DCVENDORINV) " _
                   & "VALUES('" _
                   & Trim(sJournalID) & "'," _
                   & iTrans & "," _
                   & iRef & "," _
                   & cFREIGHT & ",'" _
                   & sFrAcct & "','" _
                   & lblPst & "','" _
                   & sVendor & "','" _
                   & Trim(txtInv) & "')"
            clsADOCon.ExecuteSql sSql
         End If
      
      End If
      MouseCursor 0
      If clsADOCon.ADOErrNum = 0 Then
         clsADOCon.CommitTrans
         SysMsg "Successfully Recorded.", True
         
         ReSetInputData
      Else
         clsADOCon.RollbackTrans
         clsADOCon.ADOErrNum = 0
         MsgBox "Could Not Recorded The " & sInv & " .", _
            vbInformation, Caption
      End If
      ManageBoxes 1
      On Error Resume Next
      cmbVnd.SetFocus
   Else
      CancelTrans
      On Error Resume Next
      txtInv.SetFocus
   End If
   Exit Sub
DiaErr1:
   sProcName = "postthismem"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub


Private Function GetDBAccounts(ApAcct As String, RevAcct As String, TxAcct As String, _
                               FrAcct As String) As Byte
   Dim AdoCdm As ADODB.Recordset
   On Error GoTo DiaErr1
      sSql = "SELECT COAPACCT,COCRREVACCT,COPJTAXACCT,COPJTFRTACCT  FROM ComnTable WHERE COREF=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, AdoCdm, ES_FORWARD)
   If bSqlRows Then
      With AdoCdm
         ApAcct = "" & Trim(!COAPACCT)
         RevAcct = "" & Trim(!COCRREVACCT)
         TxAcct = "" & Trim(!COPJTAXACCT)
         FrAcct = "" & Trim(!COPJTFRTACCT)
         .Cancel
      End With
   Else
      ApAcct = ""
      RevAcct = ""
      TxAcct = ""
      FrAcct = ""
   End If
   Set AdoCdm = Nothing
   Exit Function
DiaErr1:
   sProcName = "getdbaccounts"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Function

Private Sub txtPst_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtPst_LostFocus()
   On Error Resume Next
   Dim dDate As Date
   txtPst = CheckDate(txtPst)
   lblPst = txtPst
   dDate = lblPst
   lblDue = Format(dDate + 30, "mm/dd/yy")
End Sub

Private Sub FillAccounts()
   Dim ADOAct As ADODB.Recordset
   Dim i As Integer
   
   On Error GoTo DiaErr1
   sSql = "Qry_FillLowAccounts"
   bSqlRows = clsADOCon.GetDataSet(sSql, ADOAct, ES_FORWARD)
   If bSqlRows Then
      With ADOAct
         Do Until .EOF
            For i = 0 To 4
                AddComboStr cmbAct(i).hWnd, "" & Trim(!GLACCTNO)
            Next i
            .MoveNext
         Loop
         .Cancel
      End With
   End If
   Set ADOAct = Nothing
'   For i = 0 To 4
'       If cmbAct(i).ListCount > 0 Then
'        cmbAct(i).ListIndex = 0
' '         cmbAct(i) = cmbAct(i).List(0)
''        FindAccount Me
'        lblDsc(0) = UpdateActDesc(cmbAct(i), lblDsc(0), True)
'        End If
'    Next i
   Exit Sub
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub



Private Sub UpdateTotals()
   Dim b As Byte
   Dim cTotal As Currency
   On Error Resume Next
   For b = 0 To 4
      If Trim(txtAcctAmt(b)) <> "" Then
         cTotal = cTotal + CCur(txtAcctAmt(b))
      End If
   Next
   
   If (Trim(txtTax) <> "") Then
      cTotal = cTotal + CCur(Trim(txtTax))
   End If
   
   If (Trim(txtFrt) <> "") Then
      cTotal = cTotal + CCur(Trim(txtFrt))
   End If
   
   lblTot = Format(cTotal, CURRENCYMASK)

   If CCur(lblTot) < 0 Then
      cmbTyp = "CM"
   Else
      cmbTyp = "DM"
   End If

'   If lblTot <> txtAmt Then
'      lblTot.ForeColor = ES_RED
'   Else
'      lblTot.ForeColor = diaAPe02a.ForeColor
'   End If
End Sub

Private Sub ReSetInputData()

   Dim b As Integer
   
   For b = 0 To 4
      If Trim(txtAcctAmt(b)) <> "" Then
         If Trim(cmbAct(b)) <> "" Then
            txtAcctAmt(b) = ""
            cmbAct(b) = ""
         End If
      End If
   Next

   txtFrt = ""
   txtTax = ""
   lblTot = ""

End Sub

