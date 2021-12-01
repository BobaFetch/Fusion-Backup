VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form diaAPe01a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Vendor Invoice"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7365
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   7365
   ShowInTaskbar   =   0   'False
   Tag             =   "DS"
   Begin VB.ComboBox txtDue 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1560
      TabIndex        =   11
      Tag             =   "4"
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   120
      TabIndex        =   22
      Top             =   2640
      Width           =   7095
   End
   Begin VB.CheckBox optPst 
      Height          =   195
      Left            =   1440
      TabIndex        =   20
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtCmt 
      Enabled         =   0   'False
      Height          =   975
      Left            =   1560
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Tag             =   "9"
      Top             =   4200
      Width           =   4335
   End
   Begin VB.CommandButton cmdItm 
      Caption         =   "&Items"
      Enabled         =   0   'False
      Height          =   315
      Left            =   6360
      TabIndex        =   18
      ToolTipText     =   "List Purchase Order Items"
      Top             =   2760
      Width           =   875
   End
   Begin VB.ComboBox txtPdt 
      Enabled         =   0   'False
      Height          =   315
      Left            =   4200
      TabIndex        =   10
      Tag             =   "4"
      Top             =   3480
      Width           =   1095
   End
   Begin VB.ComboBox txtIdt 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1560
      TabIndex        =   9
      Tag             =   "4"
      Top             =   3480
      Width           =   1095
   End
   Begin VB.TextBox txtAmt 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   1560
      TabIndex        =   8
      Tag             =   "1"
      Top             =   3120
      Width           =   1095
   End
   Begin VB.TextBox txtInv 
      Height          =   285
      Left            =   1560
      TabIndex        =   7
      Tag             =   "3"
      Top             =   2760
      Width           =   2775
   End
   Begin VB.ComboBox cmbPon 
      Height          =   315
      Left            =   1560
      TabIndex        =   5
      Tag             =   "2"
      ToolTipText     =   "Select PO For This Vendor From List"
      Top             =   2160
      Width           =   1095
   End
   Begin VB.ComboBox cmbVnd 
      Height          =   315
      Left            =   1560
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Contains Vendors With PO's Not Invoiced"
      Top             =   480
      Width           =   1555
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   6360
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6840
      Top             =   960
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   5400
      FormDesignWidth =   7365
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
      GroupAllowAllUp =   -1  'True
      Autosize        =   2
      RoundedCorners  =   0   'False
      BevelWidth      =   0
      Outline         =   0   'False
      PictureUp       =   "diaAPe01a.frx":0000
      PictureDn       =   "diaAPe01a.frx":0146
   End
   Begin VB.Label lblPOAmt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3240
      TabIndex        =   25
      Top             =   3120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Remit To"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   24
      Top             =   840
      Width           =   1185
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Due Date"
      Height          =   285
      Index           =   12
      Left            =   240
      TabIndex        =   23
      Top             =   3840
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice Comments:"
      Height          =   525
      Index           =   11
      Left            =   240
      TabIndex        =   19
      Top             =   4200
      Width           =   1185
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Posting Date"
      Height          =   285
      Index           =   7
      Left            =   3120
      TabIndex        =   17
      Top             =   3480
      Width           =   1065
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice Date"
      Height          =   285
      Index           =   6
      Left            =   240
      TabIndex        =   16
      Top             =   3480
      Width           =   1905
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice Amount"
      Height          =   285
      Index           =   5
      Left            =   240
      TabIndex        =   15
      Top             =   3120
      Width           =   1905
   End
   Begin VB.Label lblCount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   4800
      TabIndex        =   14
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Orders Found"
      Height          =   285
      Index           =   4
      Left            =   2880
      TabIndex        =   13
      Top             =   2160
      Width           =   2025
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice Number"
      Height          =   285
      Index           =   3
      Left            =   240
      TabIndex        =   6
      Top             =   2760
      Width           =   1905
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Orders"
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   4
      Top             =   2160
      Width           =   1905
   End
   Begin VB.Label lblNme 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   1155
      Left            =   1560
      TabIndex        =   3
      Top             =   840
      Width           =   2775
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Vendor"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   1185
   End
End
Attribute VB_Name = "diaAPe01a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of                     ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***

'See the UpdateTables prodecure for database revisions

Option Explicit

'*************************************************************************************
'
' diaAPe01a - Create AP Invoice
'
' Notes:
'
' Created: (cjs)
' Revisions:
' 08/02/01 (nth) Add the abliblity to add items not found on invoice.
' 08/07/01 (nth) Fixed error with diapspini showing even if no items not on
'                PO where present.
' 11/06/02 (nth) Add code to reset form after invoice was successfully posted.
' 12/26/02 (nth) Increased invoice size to 20 chars.
' 12/27/02 (nth) Added due date per JLH.
' 04/06/04 (nth) Check for valid PO number per JEVINT.
'
'*************************************************************************************

Dim rdoQry As ADODB.Command
Dim AdoParameter1 As ADODB.Parameter

Dim bOnLoad As Byte
Dim bCancel As Byte
Dim bGoodVendor As Byte
Dim dToday As Date
Dim iNetDays As Integer
Dim sOldVendor As String
Dim sMsg As String
Dim cPOAmt As Currency

Public sJournalID As String ' Passed to diaAPe01b

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'*************************************************************************************

Private Sub cmbPon_LostFocus()
   Dim i As Integer
   Dim bFound As Byte
   If Len(cmbPon) = 0 Then
      cmbPon = "NO PO"
      cPOAmt = 0
      lblPOAmt = Format(cPOAmt, CURRENCYMASK)
   Else
      If UCase(cmbPon) <> "NO PO" Then
         ' Check if PO can be found ?
         cmbPon = Format(cmbPon, "000000")
         For i = 0 To cmbPon.ListCount
            If cmbPon = cmbPon.List(i) Then
               bFound = True
               Exit For
            End If
         Next
         If Not bFound Then
            sMsg = "PO With Received Items Not Found."
            MsgBox sMsg, vbInformation, Caption
            cmbPon = "NO PO"
            Exit Sub
         End If
         cmbPon = Compress(cmbPon, 6)
         GetPOTerms cmbPon, iNetDays
         txtIdt_Change
                
                 cPOAmt = 0
         ' Get PO Total
         GetPOTotal cmbVnd.Text, cmbPon.Text
         If (cPOAmt = -1) Then
            lblPOAmt = Format(0, CURRENCYMASK)
         Else
            lblPOAmt = Format(cPOAmt, CURRENCYMASK)
         End If
      End If
   End If
   If bGoodVendor Then ManageBoxes True Else ManageBoxes False
End Sub

Private Sub cmbVnd_Click()
   cmbVnd = CheckLen(cmbVnd, 10)
   If cmbVnd <> sOldVendor Then
      bGoodVendor = FindVendor(Me, , iNetDays, True)
      If bGoodVendor Then GetPurchaseOrders
   End If
End Sub

Private Sub cmbVnd_LostFocus()
   cmbVnd = CheckLen(cmbVnd, 10)
   bGoodVendor = FindVendor(Me, , iNetDays, True)
   If bGoodVendor Then
          cPOAmt = 0
      GetPurchaseOrders
      txtIdt_Change
   End If
End Sub

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      SelectHelpTopic Me, "Vendor Invoice (Single PO)"
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub

Private Sub cmdItm_Click()
   Dim bBadInvoice As Byte
   Dim b As Byte
   
   'make sure dates are formatted correctly
   txtIdt = CheckDate(txtIdt)
   txtPdt = CheckDate(txtPdt)
   txtDue = CheckDate(txtDue)
   
   If DateDiff("d", CDate(txtPdt), Now) < 0 Then
        If MsgBox("You are posting an invoice to a future date." & vbCrLf & "Is this what you want to do?", vbYesNo) <> vbYes Then
            txtPdt.SetFocus
            Exit Sub
        End If
        
   End If

   
   sJournalID = GetOpenJournal("PJ", Format(txtPdt, "mm/dd/yy"))
   If Left(sJournalID, 4) = "None" Then
      sJournalID = ""
      b = 1
   Else
      If sJournalID = "" Then b = 0 Else b = 1
   End If
   
   If b = 0 Then
      MsgBox "There Is No Open Purchases Journal For The Period.", _
         vbInformation, Caption
      txtPdt.SetFocus
      Exit Sub
   End If
   
   bBadInvoice = CheckInvoice
   
   If bBadInvoice Then
      If Len(Trim(txtInv)) = 0 Then
         MsgBox "Requires An Invoice Number.", vbInformation, Caption
         On Error Resume Next
         txtInv.SetFocus
         Exit Sub
      End If
      
      'make sure all dates are in mm/dd/yy format
      
      
      If UCase(cmbPon) = "NO PO" Then
         diaAPe01b.Caption = "Vendor Invoice (No PO)"
         diaAPe01b.lblPon.Visible = False
         diaAPe01b.lblRel.Visible = False
         diaAPe01b.z1(4).Visible = False
         diaAPe01b.z1(14).Visible = False
         diaAPe01b.z1(6).Visible = False
         diaAPe01b.z1(11).Visible = False
         diaAPe01b.z1(5) = "Description/GL Account                  "
      End If
      diaAPe01b.lblPostDate = Format(txtPdt, "mm/dd/yy")
      diaAPe01b.Show
   Else
      MsgBox "That Vendor/Invoice Combination Is In Use.", _
         vbInformation, Caption
   End If
End Sub



Private Sub Form_Activate()
   Dim b As Byte
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      sJournalID = GetOpenJournal("PJ", Format(ES_SYSDATE, "mm/dd/yy"))
      'CurrentJournal "PJ", Format(ES_SYSDATE, "mm/dd/yy"), sJournalID
      If Left(sJournalID, 4) = "None" Then
         sJournalID = ""
         b = 1
      Else
         If sJournalID = "" Then b = 0 Else b = 1
      End If
      If b = 0 Then
         MsgBox "There Is No Open Purchases Journal For The Period.", _
            vbInformation, Caption
         Sleep 500
         MouseCursor 0
         Unload Me
         Exit Sub
      Else
         FillVendors Me
         FillCombo
         
         cmbVnd = cUR.CurrentVendor
         FindVendor Me, , iNetDays, True
         txtIdt_Change
         GetPurchaseOrders
         bOnLoad = False
      End If
   Else
      If optPst.Value = vbChecked Then
         ManageBoxes False
         txtInv = ""
         txtAmt = "0.00"
         If UCase(cmbPon) <> "NO PO" Then GetPurchaseOrders
         txtCmt = ""
         optPst.Value = vbUnchecked
      End If
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   sCurrForm = Caption
   sSql = "SELECT DISTINCT PONUMBER,POVENDOR,PINUMBER," _
          & "PITYPE FROM PohdTable,PoitTable WHERE " _
          & "PONUMBER=PINUMBER AND (PITYPE=15 AND " _
          & "POVENDOR= ? ) "
   Set rdoQry = New ADODB.Command
   rdoQry.CommandText = sSql
   
   Set AdoParameter1 = New ADODB.Parameter
   AdoParameter1.Type = adChar
   AdoParameter1.SIZE = 10
   rdoQry.parameters.Append AdoParameter1
   
   txtIdt = Format(ES_SYSDATE, "mm/dd/yy")
   txtPdt = Format(ES_SYSDATE, "mm/dd/yy")
   txtAmt = "0.00"
   dToday = Format(ES_SYSDATE, "mm/dd/yy")
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
   Set AdoParameter1 = Nothing
   Set rdoQry = Nothing
   Set diaAPe01a = Nothing
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Private Sub FillCombo()
   On Error GoTo DiaErr1
   If cmbVnd.ListCount > 0 Then
      bGoodVendor = FindVendor(Me, , iNetDays)
   Else
      MsgBox "There Are No Vendors.", vbInformation, Caption
   End If
   If bGoodVendor Then GetPurchaseOrders
   Exit Sub
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub GetPurchaseOrders()
   Dim RdoPon As ADODB.Recordset
   Dim iTotal As Integer
   Dim sVendor As String
   
   On Error GoTo DiaErr1
   cmbPon.Clear
   sVendor = cmbVnd
   rdoQry.parameters(0).Value = Compress(sVendor)
   bSqlRows = clsADOCon.GetQuerySet(RdoPon, rdoQry)
   If bSqlRows Then
      With RdoPon
         Do Until .EOF
            iTotal = iTotal + 1
            cmbPon.AddItem "" & Format(!PONumber, "000000")
            .MoveNext
         Loop
         .Cancel
      End With
      sOldVendor = cmbVnd
   End If
   cmbPon.AddItem "No PO"
   cmbPon = cmbPon.List(0)
   lblCount = iTotal
   
   ' Get PO Total
   GetPOTotal cmbVnd.Text, cmbPon.Text
   If (cPOAmt = -1) Then
      lblPOAmt = Format(0, CURRENCYMASK)
   Else
      lblPOAmt = Format(cPOAmt, CURRENCYMASK)
   End If
   'lblPOAmt = Format(cPOAmt, CURRENCYMASK)
     
   Set RdoPon = Nothing
   Exit Sub
DiaErr1:
   sProcName = "getpurchaseo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub txtAmt_LostFocus()
   'txtAmt = CheckLen(txtAmt, 8)
   'txtAmt = Format(Val(txtAmt), "####0.00")
   '    Dim s As String
   '    s = CheckCurrency(txtAmt, False)
   '    If s = "*" Then
   '        txtAmt.SetFocus
   '    Else
   '        txtAmt = s
   '    End If
   
   CheckCurrencyTextBox txtAmt, False
'Not needed
'   If (cPOAmt <> -1) Then
'      If (Val(cPOAmt) <> Val(txtAmt)) Then
'         MsgBox "The PO Amount is not same as the Invoice Amount." & vbCrLf _
'               & "The PO Amount =" & Format(cPOAmt, CURRENCYMASK) & ".", _
'                     vbInformation, Caption
'      End If
'   End If
End Sub

Private Sub txtCmt_LostFocus()
   txtCmt = CheckLen(txtCmt, 1020)
   txtCmt = CheckComments(txtCmt)
End Sub

Private Sub txtDue_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtDue_LostFocus()
   txtDue = CheckDate(txtDue)
End Sub

Private Sub txtIdt_Change()
   On Error Resume Next
   Dim dDate As Date
   dDate = txtIdt
   txtDue = Format(DateAdd("d", iNetDays, dDate), "mm/dd/yy")
End Sub

Private Sub txtIdt_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtIdt_LostFocus()
   txtIdt = CheckDate(txtIdt)
End Sub

Private Sub txtInv_Click()
   bGoodVendor = FindVendor(Me, , iNetDays)
   If bGoodVendor Then ManageBoxes True
End Sub

Private Sub txtInv_LostFocus()
   txtInv = CheckComments(CheckLen(txtInv, 20))
End Sub

Private Sub txtPdt_DropDown()
   ShowCalendar Me
End Sub

Private Sub ManageBoxes(bOpen As Boolean)
   On Error Resume Next
   If bOpen Then
      cmdItm.enabled = True
      txtAmt.enabled = True
      txtIdt.enabled = True
      txtPdt.enabled = True
      txtDue.enabled = True
      txtCmt.enabled = True
   Else
      cmdItm.enabled = False
      txtAmt.enabled = False
      txtIdt.enabled = False
      txtPdt.enabled = False
      txtDue.enabled = False
      txtCmt.enabled = False
   End If
End Sub

Private Function CheckInvoice() As Byte
   Dim rdoOld As ADODB.Recordset
   Dim sVendor As String
   sVendor = Compress(cmbVnd)
   On Error GoTo DiaErr1
   sSql = "SELECT VINO,VIVENDOR FROM VihdTable WHERE " _
          & "VINO='" & txtInv & "' AND VIVENDOR='" & sVendor & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoOld)
   If bSqlRows Then
      CheckInvoice = 0
   Else
      CheckInvoice = 1
   End If
   On Error Resume Next
   rdoOld.Close
   Set rdoOld = Nothing
   Exit Function
DiaErr1:
   sProcName = "checkinvoice"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Function

Private Sub GetPOTerms(lPO As Long, Optional ByRef iPONetDays As Integer)
   Dim rdoPo As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT ISNULL(PONETDAYS, 0) PONETDAYS FROM PohdTable WHERE PONUMBER = " & CStr(lPO)
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoPo)
   If bSqlRows Then
      iPONetDays = rdoPo!PONETDAYS
   End If
   Set rdoPo = Nothing
   Exit Sub
DiaErr1:
   sProcName = "GetPOTerms"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub


Private Sub GetPOTotal(strVendor As String, strPONum As String)
   Dim rdoPo As ADODB.Recordset
   On Error GoTo DiaErr1
   
   If ((strVendor <> "") And (strPONum <> "") And IsNumeric(strPONum)) Then
      sSql = "SELECT ISNULL(SUM(PIAMT + PIADDERS), 0) POTot FROM PohdTable, PoitTable WHERE " _
             & "PONUMBER=PINUMBER AND PINUMBER= '" & strPONum & "' AND PITYPE=15 AND POVENDOR ='" & Compress(strVendor) & "'" _
      
      bSqlRows = clsADOCon.GetDataSet(sSql, rdoPo)
      If bSqlRows Then
         cPOAmt = rdoPo!POTot
      End If
      Set rdoPo = Nothing
   Else
      cPOAmt = -1
   End If
   
   Exit Sub
DiaErr1:
   sProcName = "GetPOTotal"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub txtPdt_LostFocus()
   txtPdt = CheckDate(txtPdt)
End Sub
