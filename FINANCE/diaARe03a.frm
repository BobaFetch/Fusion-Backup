VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form diaARe03a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Customer Debit Or Credit Memo"
   ClientHeight    =   8640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7080
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   7080
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtReasons 
      Height          =   975
      Left            =   1800
      MultiLine       =   -1  'True
      TabIndex        =   53
      Tag             =   "9"
      Top             =   7200
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.ComboBox cboFedTax 
      Height          =   315
      Left            =   1800
      Sorted          =   -1  'True
      TabIndex        =   13
      Tag             =   "3"
      Top             =   5160
      Width           =   1455
   End
   Begin VB.TextBox txtFedTaxRate 
      Height          =   285
      Left            =   1800
      TabIndex        =   9
      Tag             =   "1"
      ToolTipText     =   "Tax Percentage (Positive Values)"
      Top             =   3720
      Width           =   615
   End
   Begin VB.ComboBox cmbPre 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1800
      TabIndex        =   4
      Tag             =   "3"
      ToolTipText     =   "Invoice Prefix(A-Z)"
      Top             =   2640
      Width           =   510
   End
   Begin VB.TextBox txtInv 
      Height          =   315
      Left            =   2280
      TabIndex        =   5
      Tag             =   "1"
      ToolTipText     =   "Requires A Number"
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton cmdComments 
      DisabledPicture =   "diaARe03a.frx":0000
      DownPicture     =   "diaARe03a.frx":0972
      Height          =   350
      Left            =   6240
      Picture         =   "diaARe03a.frx":12E4
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Standard Comments"
      Top             =   6000
      Width           =   350
   End
   Begin VB.TextBox txtStAdr 
      Height          =   645
      Left            =   1800
      MultiLine       =   -1  'True
      TabIndex        =   2
      Tag             =   "9"
      Top             =   1440
      Width           =   3375
   End
   Begin VB.ComboBox txtDte 
      Height          =   315
      Left            =   4680
      TabIndex        =   6
      Tag             =   "4"
      Top             =   2640
      Width           =   1095
   End
   Begin VB.ComboBox txtReva 
      Height          =   315
      Left            =   1800
      Sorted          =   -1  'True
      TabIndex        =   14
      Tag             =   "3"
      Top             =   5520
      Width           =   1455
   End
   Begin VB.ComboBox txtTxa 
      Height          =   315
      Left            =   1800
      Sorted          =   -1  'True
      TabIndex        =   12
      Tag             =   "3"
      Top             =   4800
      Width           =   1455
   End
   Begin VB.ComboBox txtFra 
      Height          =   315
      Left            =   1800
      Sorted          =   -1  'True
      TabIndex        =   11
      Tag             =   "3"
      Top             =   4440
      Width           =   1455
   End
   Begin VB.TextBox txtCmt 
      Height          =   975
      Left            =   1800
      MultiLine       =   -1  'True
      TabIndex        =   15
      Tag             =   "9"
      Top             =   6000
      Width           =   4335
   End
   Begin VB.TextBox txtFrt 
      Height          =   285
      Left            =   1800
      TabIndex        =   10
      Tag             =   "1"
      ToolTipText     =   " Freight (Positive Values)"
      Top             =   4080
      Width           =   975
   End
   Begin VB.TextBox txtTax 
      Height          =   285
      Left            =   1800
      TabIndex        =   8
      Tag             =   "1"
      ToolTipText     =   "Tax Percentage (Positive Values)"
      Top             =   3360
      Width           =   615
   End
   Begin VB.TextBox txtTot 
      Height          =   285
      Left            =   1800
      TabIndex        =   7
      Tag             =   "1"
      ToolTipText     =   "Total Except Tax And Freight (Positive Values)"
      Top             =   3000
      Width           =   1095
   End
   Begin VB.ComboBox cmbSlp 
      Height          =   315
      Left            =   1800
      Sorted          =   -1  'True
      TabIndex        =   3
      Tag             =   "3"
      ToolTipText     =   "Select Salesperson From List"
      Top             =   2160
      Width           =   975
   End
   Begin VB.ComboBox cmbCst 
      Height          =   315
      Left            =   1800
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   720
      Width           =   1555
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "P&ost"
      Enabled         =   0   'False
      Height          =   315
      Left            =   6120
      TabIndex        =   17
      ToolTipText     =   "Add This Packing Slip Invoice"
      Top             =   600
      Width           =   875
   End
   Begin VB.ComboBox cmbTyp 
      ForeColor       =   &H00800000&
      Height          =   315
      ItemData        =   "diaARe03a.frx":1C56
      Left            =   1800
      List            =   "diaARe03a.frx":1C58
      TabIndex        =   0
      Tag             =   "8"
      ToolTipText     =   "Select CM (Credit Memo) Or DM (Debit Memo) From List"
      Top             =   360
      Width           =   735
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   6120
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   19
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
      PictureUp       =   "diaARe03a.frx":1C5A
      PictureDn       =   "diaARe03a.frx":1DA0
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6480
      Top             =   3600
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   8640
      FormDesignWidth =   7080
   End
   Begin Threed.SSRibbon ShowPrinters 
      Height          =   225
      Left            =   360
      TabIndex        =   42
      ToolTipText     =   "Show System Printers"
      Top             =   0
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   397
      _StockProps     =   65
      BackColor       =   12632256
      GroupAllowAllUp =   -1  'True
      RoundedCorners  =   0   'False
      BevelWidth      =   0
      Outline         =   0   'False
      PictureUp       =   "diaARe03a.frx":1EE6
      PictureDn       =   "diaARe03a.frx":202C
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Reasons"
      Height          =   285
      Index           =   20
      Left            =   120
      TabIndex        =   54
      Top             =   7200
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Misc FederalTax"
      Height          =   285
      Index           =   19
      Left            =   120
      TabIndex        =   52
      Top             =   5160
      Width           =   1320
   End
   Begin VB.Label lblFedTax 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3360
      TabIndex        =   51
      Top             =   5160
      Width           =   2775
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Fed Tax Adjustment"
      Height          =   285
      Index           =   18
      Left            =   120
      TabIndex        =   50
      Top             =   3720
      Width           =   1800
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      Height          =   285
      Index           =   17
      Left            =   2520
      TabIndex        =   49
      Top             =   3720
      Width           =   345
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Federal Tax"
      Height          =   285
      Index           =   1
      Left            =   3600
      TabIndex        =   48
      Top             =   3720
      Width           =   1035
   End
   Begin VB.Label lblFedTaxAmount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      Height          =   285
      Left            =   4680
      TabIndex        =   47
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label lblJrn 
      Height          =   255
      Left            =   6120
      TabIndex        =   46
      Top             =   1560
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblDummy 
      Caption         =   "dummy"
      Height          =   255
      Left            =   6120
      TabIndex        =   45
      Top             =   1920
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice Number"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   44
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label lblPrinter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Default Printer"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   675
      TabIndex        =   43
      Top             =   0
      Width           =   2760
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ship To:"
      Height          =   255
      Index           =   15
      Left            =   120
      TabIndex        =   41
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Posting Date"
      Height          =   255
      Index           =   14
      Left            =   3600
      TabIndex        =   40
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label lblReva 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3360
      TabIndex        =   39
      Top             =   5520
      Width           =   2775
   End
   Begin VB.Label lblTxa 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3360
      TabIndex        =   38
      Top             =   4800
      Width           =   2775
   End
   Begin VB.Label lblFra 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3360
      TabIndex        =   37
      Top             =   4440
      Width           =   2775
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Comments"
      Height          =   285
      Index           =   13
      Left            =   120
      TabIndex        =   36
      Top             =   6000
      Width           =   1320
   End
   Begin VB.Label lblTax 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      Height          =   285
      Left            =   4680
      TabIndex        =   35
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Tax"
      Height          =   285
      Index           =   10
      Left            =   3600
      TabIndex        =   34
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label lblTot 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      Height          =   285
      Left            =   4680
      TabIndex        =   33
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      Height          =   285
      Index           =   12
      Left            =   3600
      TabIndex        =   32
      Top             =   4080
      Width           =   1185
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      Height          =   285
      Index           =   11
      Left            =   2520
      TabIndex        =   31
      Top             =   3360
      Width           =   345
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Misc Revenue"
      Height          =   285
      Index           =   9
      Left            =   120
      TabIndex        =   30
      Top             =   5520
      Width           =   1320
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Misc Sales Tax"
      Height          =   285
      Index           =   8
      Left            =   120
      TabIndex        =   29
      Top             =   4800
      Width           =   1320
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Misc Freight"
      Height          =   285
      Index           =   7
      Left            =   120
      TabIndex        =   28
      Top             =   4440
      Width           =   1320
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      Height          =   285
      Index           =   6
      Left            =   120
      TabIndex        =   27
      Top             =   3000
      Width           =   1320
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Tax Adjustment"
      Height          =   285
      Index           =   5
      Left            =   120
      TabIndex        =   26
      Top             =   3360
      Width           =   1800
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Freight Adjustment"
      Height          =   285
      Index           =   4
      Left            =   120
      TabIndex        =   25
      Top             =   4080
      Width           =   1320
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Salesperson"
      Height          =   285
      Index           =   16
      Left            =   120
      TabIndex        =   24
      Top             =   2160
      Width           =   1545
   End
   Begin VB.Label lblSlp 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2880
      TabIndex        =   23
      Top             =   2160
      Width           =   2775
   End
   Begin VB.Label txtNme 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1800
      TabIndex        =   22
      Top             =   1095
      Width           =   2775
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   21
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice Type"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   20
      Top             =   360
      Width           =   1575
   End
End
Attribute VB_Name = "diaARe03a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of                     ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***

'See the UpdateTables prodecure for database revisions


'************************************************************************************
' diaARe03a - Credit of Debit memo.
'
' Created: (cjs)
' Modified:
'   ??/??/?? (nth) To require invoice adjustment amount before posting
'   07/10/01 (nth) Invoice total stored in the database is to include both freight and tax.
'   07/16/01 (nth) Add ship-to address to CM and DM
'   08/14/02 (nth) Allow revision of ship-to address
'   08/30/02 (nth) Fixed error with DM amount showing negative
'   06/04/03 (nth) Fixed check for sales journal
'   08/15/03 (nth) Fixed invtotal doubling tax on CM per WCK
'   08/15/03 (nth) Calculate sales tax from customer tax code per WCK
'   10/28/04 (nth) To record rev account in invoice header see INVCRACCT.
'
'************************************************************************************

Option Explicit

Dim bOnLoad As Byte
Dim bCancel As Byte
Dim bGoodCustomer As Byte
Dim bGlverify As Byte
Dim bInvUsed As Boolean

Dim lNewInv As Long
Dim sInvPre As String
Dim sJournalID As String
Dim sMsg As String

Dim sCOSjARAcct As String
Dim sCOSjFrtAcct As String
Dim sCOSjTaxAcct As String
Dim sCOCrRevAcct As String
Dim sCOSjFedTaxAcct As String

' federal tax rate
' automatically applied if tax is being calculated per line item
Dim taxPerLineItem As Boolean
Dim fedTaxRate As Currency

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'************************************************************************************

Private Sub cmbCst_Click()
   FindCustomer Me, cmbCst
   If Len(txtNme) = 0 Then
      txtNme = "*** Invalid Customer ***"
   Else
      GetShipTo cmbCst
      GetSalesTax cmbCst, txtTax
      
   End If
End Sub

Private Sub cmbCst_LostFocus()
   cmbCst = CheckLen(cmbCst, 10)
   FindCustomer Me, cmbCst
   If Len(txtNme) = 0 Then
      txtNme = "*** Invalid Customer ***"
   Else
      GetShipTo cmbCst
      GetSalesTax cmbCst, txtTax
   End If
End Sub

Private Sub cmbPre_LostFocus()
   cmbPre = CheckLen(cmbPre, 1)
   If Asc(cmbPre) < 65 Or Asc(cmbPre) > 90 Then
      Beep
      cmbPre = sInvPre
   End If
End Sub

Private Sub cmbSlp_Click()
   GetSalesPerson
End Sub

Private Sub cmbSlp_LostFocus()
   Dim b As Byte
   Dim i As Integer
   cmbSlp = CheckLen(cmbSlp, 4)
   If Trim(cmbSlp) = "" Then
      If cmbSlp.ListCount > 0 Then
         Beep
         cmbSlp = cmbSlp.List(0)
      End If
   End If
   b = 0
   If cmbSlp.ListCount > 0 Then
      For i = 0 To cmbSlp.ListCount - 1
         If Trim(cmbSlp) = Trim(cmbSlp.List(i)) Then b = 1
      Next
      If b = 0 Then
         Beep
         cmbSlp = cmbSlp.List(0)
      End If
      GetSalesPerson
   End If
End Sub

Private Sub cmbTyp_Click()
   If cmbTyp = "CM" Then z1(12) = "CM Total" Else _
               z1(12) = "DM Total"
End Sub

Private Sub cmbTyp_LostFocus()
   cmbTyp = CheckLen(cmbTyp, 2)
   If cmbTyp <> "CM" And cmbTyp <> "DM" Then
      Beep
      cmbTyp = "CM"
   End If
   If cmbTyp = "CM" Then z1(12) = "CM Total" Else _
               z1(12) = "DM Total"
End Sub


Private Sub cmdAdd_Click()
   ' Make sure invoice has a total before allowing invoice to post
   If CCur(lblTot) = 0 Then
      sMsg = "Invoice Adjusted Amount Is Required"
      MsgBox sMsg, vbInformation, Caption
      Exit Sub
   End If
   
   If Trim(txtReva) = "" Then
      sMsg = "Revenue Account Is Required."
      MsgBox sMsg, vbInformation, Caption
      txtReva.SetFocus
      Exit Sub
   End If
   
   If CSng(txtTax) > 0 And Trim(txtTxa) = "" Then
      sMsg = "Tax Account Is Required."
      MsgBox sMsg, vbInformation, Caption
      txtTax.SetFocus
      Exit Sub
   End If
   
   If CCur(txtFedTaxRate) > 0 And Trim(cboFedTax) = "" Then
      sMsg = "Tax Account Is Required."
      MsgBox sMsg, vbInformation, Caption
      txtFedTaxRate.SetFocus
      Exit Sub
   End If
   
   If CCur(txtFrt) > 0 And Trim(txtFra) = "" Then
      sMsg = "Freight Account Is Required."
      MsgBox sMsg, vbInformation, Caption
      txtFrt.SetFocus
      Exit Sub
   End If
   
   ' Check if invoice number is used.
   If Val(txtInv) <> lNewInv Then
      bInvUsed = GetOldInvoice(txtInv)
   Else
      bInvUsed = False
   End If
   
   If bInvUsed Then
      MsgBox "Invoice Number Is In Use.", vbInformation, Caption
      txtInv = Format(lNewInv, "000000")
   Else
      UpdateInvoice
   End If
   
End Sub

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      SelectHelpTopic Me, "Customer Debit Or Credit Memo"
      cmdHlp = False
      MouseCursor 0
   End If
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      CurrentJournal "SJ", ES_SYSDATE, sJournalID
      lblJrn.ForeColor = ES_RED
      lblJrn = "One Or More Required Accounts Is Missing."
      FillCustomers Me
      If cUR.CurrentCustomer <> "" Then
         cmbCst = cUR.CurrentCustomer
      Else
         If cmbCst.ListCount > 0 Then cmbCst = cmbCst.List(0)
      End If
      FindCustomer Me, cmbCst
      Dim nTaxRate As Single
      GetSalesTax cmbCst, nTaxRate
      txtTax = Format(nTaxRate, "0.000")
      FillSalesPersons
      FillAccounts
      AddInvoice
      FindCustomer Me, cmbCst
      If Len(txtNme) = 0 Then
         txtNme = "*** Invalid Customer ***"
      Else
         GetShipTo cmbCst
      End If
      
      ' is tax calculated per line item?
      taxPerLineItem = IsTaxCalculatedPerLineItem()
      
      bOnLoad = False
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   Dim i As Integer
   FormLoad Me
   FormatControls
   sCurrForm = Caption
   sInvPre = GetSetting("Esi2000", "EsiFina", "LastInvPref", sInvPre)
   cmbTyp.AddItem "CM"
   cmbTyp.AddItem "DM"
   cmbTyp = "CM"
   If Len(Trim(sInvPre)) = 0 Then sInvPre = "I"
   cmbPre = sInvPre
   For i = 65 To 90
      AddComboStr cmbPre.hWnd, Chr(i)
   Next
   lblFra.ForeColor = vbBlack
   lblTxa.ForeColor = vbBlack
   lblReva.ForeColor = vbBlack
   
   CloseBoxes
   GetOptions
   bOnLoad = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Dim bResponse As Byte
   Dim sMsg As String
   If Val(lblTot) > 0 Then
      sMsg = "Do You Really Want To Cancel The Addition" & vbCrLf _
             & "Of Memo " & txtInv & " ?"
      bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
      If bResponse = vbNo Then
         Cancel = True
      End If
   End If
   On Error Resume Next
   'dump any left over dummies
   sSql = "DELETE FROM CihdTable WHERE INVNO=" & lNewInv _
          & "AND INVTYPE='TM'"
   clsADOCon.ExecuteSQL sSql
   
   'Save Accounts
   SaveSetting "Esi2000", "EsiFina", "LastInvPref", cmbPre
   SaveSetting "Esi2000", "EsiFina", "MemotxtFra", sCOSjFrtAcct
   SaveSetting "Esi2000", "EsiFina", "MemoTxtTxa", sCOSjTaxAcct
   SaveSetting "Esi2000", "EsiFina", "MemoTxtRev", sCOCrRevAcct
   SaveSetting "Esi2000", "EsiFina", "MemoTxtRev", sCOCrRevAcct
   SaveSetting "Esi2000", "EsiFina", "MemoFedTaxRate", CStr(fedTaxRate)
   If Len(cmbCst) Then
      cUR.CurrentCustomer = cmbCst
      SaveCurrentSelections
   End If
   SaveOptions
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set diaARe03a = Nothing
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtDte = Format(Now, "mm/dd/yy")
End Sub

Private Sub AddInvoice()
   
   GetNextInvoice
   
   On Error GoTo DiaErr1
   
   ' Reserve a record in case the invoice number is changed.
   ' Use TM so that it won't show and can be safely deleted.
   
   sSql = "INSERT INTO CihdTable (INVNO,INVTYPE) " _
          & "VALUES(" & lNewInv & ",'TM')"
   clsADOCon.ExecuteSQL sSql
   Exit Sub
   
DiaErr1:
   sProcName = "addinvoice"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub lblFra_Change()
   If Left(lblFra, 6) = "*** Ac" Then
      lblFra.ForeColor = ES_RED
   Else
      lblFra.ForeColor = vbBlack
   End If
   
End Sub

Private Sub lblReva_Change()
   If Left(lblReva, 6) = "*** Ac" Then
      lblReva.ForeColor = ES_RED
   Else
      lblReva.ForeColor = vbBlack
   End If
   
End Sub

Private Sub lblTax_Change()
   If Val(lblTax) > Val(txtTot) Then
      lblTax.ForeColor = ES_RED
      cmdAdd.enabled = False
   Else
      lblTax.ForeColor = vbBlack
      cmdAdd.enabled = True
   End If
   
End Sub

Private Sub lblTxa_Change()
   If Left(lblTxa, 6) = "*** Ac" Then
      lblTxa.ForeColor = ES_RED
   Else
      lblTxa.ForeColor = vbBlack
   End If
End Sub

Private Sub ShowPrinters_Click(Value As Integer)
   SysPrinters.Show
   ShowPrinters.Value = False
End Sub

Private Sub txtCmt_LostFocus()
   txtCmt = CheckLen(txtCmt, 2048)
End Sub


Private Sub txtDte_DropDown()
   ShowCalendar Me
End Sub


Private Sub txtDte_LostFocus()
   txtDte = CheckDate(txtDte)
End Sub


Private Sub txtFedTaxRate_LostFocus()
   txtFedTaxRate = CheckLen(txtFedTaxRate, 6)
   txtFedTaxRate = Format(Abs(Val(txtFedTaxRate)), "#0.000")
   If txtFedTaxRate > 30 Then
      Beep
      txtFedTaxRate = "0.000"
   End If
   fedTaxRate = CCur(txtFedTaxRate.Text)
   UpdateTotals
End Sub

Private Sub txtFra_Click()
   GetAccount "txtfra", txtFra
End Sub

Private Sub txtFra_LostFocus()
   txtFra = CheckLen(txtFra, 12)
   GetAccount "txtfra", txtFra
   If lblFra.ForeColor = vbBlack Then sCOSjFrtAcct = Compress(txtFra)
   
End Sub


Private Sub txtFrt_LostFocus()
   txtFrt = CheckLen(txtFrt, 9)
   txtFrt = Format(Abs(Val(txtFrt)), "#####0.00")
   UpdateTotals
   
End Sub


Private Sub txtInv_LostFocus()
   txtInv = CheckLen(txtInv, 6)
   On Error Resume Next
   '    If Val(txtInv) < lNewInv Then
   '        Beep
   '        txtInv = lNewInv
   '    End If
   
   'make sure invoice number not in use
   Dim InvNo As Long
   InvNo = Val("0" & txtInv)
   If InvNo = 0 Then
      MsgBox "Invoice number required", vbInformation, "Invoice number"
      Exit Sub
   End If
   
   If InvNo <> lNewInv Then
      If Not IsInvoiceNumberAvailable(InvNo) Then
         txtInv = lNewInv
      End If
   End If
   
   txtInv = Format(Abs(Val(txtInv)), "000000")
   
   clsADOCon.ADOErrNum = 0
   
   If Val(txtInv) <> lNewInv Then
      sSql = "UPDATE CihdTable SET INVNO=" & Val(txtInv) & " " _
             & "WHERE INVNO=" & lNewInv & " AND INVTYPE='TM'"
      clsADOCon.ExecuteSQL sSql
      If clsADOCon.ADOErrNum = 0 Then
         lNewInv = Val(txtInv)
      Else
         txtInv = Format(lNewInv, "000000")
      End If
   End If
   
End Sub


Private Sub txtNme_Change()
   If Left(txtNme, 3) = "***" Then
      txtNme.ForeColor = ES_RED
      bGoodCustomer = False
   Else
      txtNme.ForeColor = vbBlack
      bGoodCustomer = True
   End If
End Sub


Private Sub FillSalesPersons()
   Dim RdoCmb As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT SPNUMBER FROM SprsTable "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCmb)
   If bSqlRows Then
      With RdoCmb
         Do Until .EOF
            AddComboStr cmbSlp.hWnd, "" & Trim(!SPNumber)
            .MoveNext
         Loop
         .Cancel
      End With
   End If
   Set RdoCmb = Nothing
   If cmbSlp.ListCount > 0 Then
      cmbSlp = cmbSlp.List(0)
      GetSalesPerson
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "findsalespe"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub GetSalesPerson()
   Dim rdoSlp As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT SPNUMBER,SPFIRST,SPLAST,SPREGION FROM SprsTable " _
          & "WHERE SPNUMBER='" & cmbSlp & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoSlp)
   If bSqlRows Then
      cmbSlp = "" & Trim(rdoSlp!SPNumber)
      lblSlp = "" & Trim(rdoSlp!SPFIRST) & " " & Trim(rdoSlp!SPLAST)
   Else
      lblSlp = ""
   End If
   On Error Resume Next
   Set rdoSlp = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getsalesper"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub GetShipTo(sCustRef As String)
   Dim RdoShp As ADODB.Recordset
   
   On Error GoTo DiaErr1
   sSql = "SELECT CUSTNAME,CUSTADR,CUSTCITY,CUSTSTATE,CUSTZIP FROM CustTable " _
          & "WHERE CUREF = '" & Compress(sCustRef) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoShp)
   
   If bSqlRows Then
      With RdoShp
         txtStAdr = "" & Trim(!CUSTNAME) & vbCrLf _
                    & Trim(!CUSTADR) & vbCrLf _
                    & Trim(!CUSTCITY) & ", " & Trim(!CUSTSTATE) & "  " & !CUSTZIP
      End With
   End If
   Set RdoShp = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "GetShipTo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub txtReva_Click()
   GetAccount "txtReva", txtReva
End Sub

Private Sub txtReva_LostFocus()
   txtReva = CheckLen(txtReva, 12)
   GetAccount "txtReva", txtReva
   If lblReva.ForeColor = vbBlack Then sCOCrRevAcct = Compress(txtReva)
End Sub

Private Sub txtStAdr_LostFocus()
   txtStAdr = CheckLen(txtStAdr, 255)
End Sub

Private Sub txtTax_LostFocus()
   txtTax = CheckLen(txtTax, 6)
   txtTax = Format(Abs(Val(txtTax)), "#0.000")
   If txtTax > 30 Then
      Beep
      txtTax = "0.000"
   End If
   UpdateTotals
End Sub

Private Sub txtTot_LostFocus()
   txtTot = CheckLen(txtTot, 9)
   txtTot = Format(Abs(Val(txtTot)), "#####0.00")
   UpdateTotals
End Sub

Private Sub UpdateTotals()
   Dim cFREIGHT As Currency
   Dim cTax As Currency
   Dim cFedTax As Currency
   Dim cTotal As Currency
   
   On Error GoTo DiaErr1
   cTotal = CCur(txtTot)
   cTax = (txtTax / 100) * cTotal
   cFedTax = (txtFedTaxRate / 100) * cTotal
   
   cFREIGHT = CCur(txtFrt)
   
   lblTax = Format(cTax, "####0.00")
   lblFedTaxAmount = Format(cFedTax, "####0.00")
   lblTot = Format(cTotal + cTax + cFedTax + cFREIGHT, "#####0.00")
   
   If bGoodCustomer Then
      If CCur(lblTot) > 0 Then
         cmdAdd.enabled = True
      Else
         cmdAdd.enabled = False
      End If
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "updatetotals"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub CloseBoxes()
   txtTot = "0.00"
   txtTax = "0.000"
   txtFedTaxRate = "0.000"
   txtFrt = "0.00"
   lblTot = "0.00"
   lblTax = "0.00"
   lblFedTaxAmount = "0.00"
   txtStAdr = ""
   txtCmt = ""
End Sub

Private Sub txtTxa_Click()
   GetAccount "txtTxa", txtTxa
End Sub

Private Sub cboFedTax_Click()
   GetAccount "cboFedTax", cboFedTax
End Sub

Private Sub txtTxa_LostFocus()
   txtTxa = CheckLen(txtTxa, 12)
   GetAccount "txtTxa", txtTxa
   If lblTxa.ForeColor = vbBlack Then sCOSjTaxAcct = Compress(txtTxa)
End Sub

Private Sub cboFedTax_LostFocus()
   cboFedTax = CheckLen(cboFedTax, 12)
   GetAccount "cboFedTax", cboFedTax
   If lblFedTax.ForeColor = vbBlack Then sCOSjFedTaxAcct = Compress(cboFedTax)
End Sub

Private Sub UpdateInvoice()
   Dim b As Byte
   Dim bResponse As Byte
   'Dim iTrans   As Integer
   'Dim iRef     As Integer
   Dim cTotal As Currency
   Dim cTax As Currency
   Dim cFedTax As Currency
   Dim cFREIGHT As Currency
   Dim sCust As String
   Dim sMsg As String
   Dim sType As String
   Dim sTemp As String
   
   On Error GoTo DiaErr1
   
   ' Get invoice type and ask to proceed.
   sType = Trim(cmbTyp)
   
   ' Check for an open journal
   sJournalID = GetOpenJournal("SJ", txtDte)
   If Left(sJournalID, 4) = "None" Then
      sJournalID = ""
      b = 1
   Else
      If sJournalID = "" Then b = 0 Else b = 1
   End If
   If b = 0 Then
      MsgBox "There Is No Open Sales Journal For The Period.", _
         vbInformation, Caption
      Exit Sub
   End If
   
   ' Every thing is good, lets post the invoice
   sCust = Compress(cmbCst)
   cTax = CCur(lblTax)
   cFedTax = CCur(lblFedTaxAmount)
   cFREIGHT = CCur(txtFrt)
   cTotal = CCur(txtTot)
   
   ' Make the values negative if invoice type is credit memo
   If sType = "CM" Then
      cTotal = cTotal * -1
      cTax = cTax * -1
      cFedTax = cFedTax * -1
      cFREIGHT = cFREIGHT * -1
   End If
   
   ' MM MsgBox ("Sales JournalAcc: " & sCOSjARAcct)
   
   If (sCOSjARAcct = "") Then
        sMsg = "Please select Account Receivable account."
        MsgBox sMsg, vbInformation, Caption
        Exit Sub
   End If
   
   If (sCOCrRevAcct = "") Then
        sMsg = "Please select Revenue Account."
        MsgBox sMsg, vbInformation, Caption
        Exit Sub
   End If
   
   
   ' Update 'TM' invoice record
   On Error Resume Next
   Err = 0
   
   clsADOCon.BeginTrans
   clsADOCon.ADOErrNum = 0
   
   sSql = "UPDATE CihdTable SET INVNO=" & lNewInv & "," _
          & "INVPRE='" & cmbPre & "'," _
          & "INVTYPE='" & sType & "'," _
          & "INVCUST='" & Compress(cmbCst) & "'," _
          & "INVTOTAL=" & (cTotal + cTax + cFREIGHT) & "," _
          & "INVTAX=" & cTax & "," _
          & "INVFEDTAX=" & cFedTax & "," _
          & "INVFEDTAXACCT='" & sCOSjFedTaxAcct & "'," _
          & "INVFREIGHT=" & cFREIGHT & "," _
          & "INVFRTACCT='" & sCOSjFrtAcct & "'," _
          & "INVTAXACCT='" & sCOSjTaxAcct & "'," _
          & "INVARACCT='" & sCOSjARAcct & "'," _
          & "INVCRACCT='" & sCOCrRevAcct & "'," _
          & "INVCOMMENTS='" & Trim(txtCmt) & "'," _
          & "INVREASONS='" & Trim(txtReasons) & "'," _
          & "INVDATE='" & Format(txtDte, "mm/dd/yy") & "'," _
          & "INVPAY=0," _
          & "INVSHIPDATE='" & Format(ES_SYSDATE, "mm/dd/yy") & "'," _
          & "INVCANCELED=0," _
          & "INVPIF=0,INVSTADR = '" & Trim(txtStAdr) & "'," _
          & "INVSLSMAN='" & cmbSlp & "' WHERE " _
          & "INVNO=" & lNewInv & " AND INVTYPE='TM'"
   clsADOCon.ExecuteSQL sSql
   
   Dim inv As New ClassARInvoice
   inv.SaveLastInvoiceNumber lNewInv
      
   ' Make journal entries
   'If sJournalID <> "" Then
   '    iTrans = GetNextTransaction(sJournalID)
   'End If
   
   Dim gl As New GLTransaction
   gl.JournalID = Trim(sJournalID)
   Dim sDate As String
   sDate = Format(txtDte, "mm/dd/yy")
   gl.InvoiceDate = CDate(sDate)
   gl.InvoiceNumber = lNewInv
   
   'If iTrans > 0 Then
   
   ' Credit memo routine...
   ' no longer need to handle debits and credits separately -- gl transaction reverses if negative
   '        If cmbTyp = "CM" Then
   '
   '            ' Credit A/R
   ''            iRef = iRef + 1
   ''            sSql = "INSERT INTO JritTable (DCHEAD,DCTRAN,DCREF," _
   ''                & "DCCREDIT,DCACCTNO,DCDATE,DCCUST,DCINVNO) VALUES('" _
   ''                & Trim(sJournalID) & "'," _
   ''                & iTrans & "," _
   ''                & iRef & "," _
   ''                & Abs((cTotal + cFREIGHT + cTax)) & ",'" _
   ''                & sCOSjARAcct & "','" _
   ''                & Format(txtDte, "mm/dd/yy") & "','" _
   ''                & sCust & "'," _
   ''                & lNewInv & ")"
   ''            clsAdoCon.ExecuteSQL sSQL
   '
   '            gl.AddDebitCredit 0, cTotal + cFREIGHT + cTax + cFedTax, sCOSjARAcct, "", 0, 0, "", sCust
   '
   '            ' Debit
   ''            iRef = iRef + 1
   ''            sSql = "INSERT INTO JritTable (DCHEAD,DCTRAN,DCREF," _
   ''                & "DCDEBIT,DCACCTNO,DCDATE,DCCUST,DCINVNO) VALUES('" _
   ''                & Trim(sJournalID) & "'," _
   ''                & iTrans & "," _
   ''                & iRef & "," _
   ''                & Abs(cTotal) & ",'" _
   ''                & sCOCrRevAcct & "','" _
   ''                & Format(txtDte, "mm/dd/yy") & "','" _
   ''                & sCust & "'," _
   ''                & lNewInv & ")"
   ''            clsAdoCon.ExecuteSQL sSQL
   '            gl.AddDebitCredit cTotal, 0, sCOCrRevAcct, "", 0, 0, "", sCust
   '
   '            ' Freight
   ''            If cFREIGHT <> 0 Then
   ''                iRef = iRef + 1
   ''                sSql = "INSERT INTO JritTable (DCHEAD,DCTRAN,DCREF," _
   ''                    & "DCDEBIT,DCACCTNO,DCDATE,DCCUST,DCINVNO) VALUES('" _
   ''                    & Trim(sJournalID) & "'," _
   ''                    & iTrans & "," _
   ''                    & iRef & "," _
   ''                    & Abs(cFREIGHT) & ",'" _
   ''                    & sCOSjFrtAcct & "','" _
   ''                    & Format(txtDte, "mm/dd/yy") & "','" _
   ''                    & sCust & "'," _
   ''                    & lNewInv & ")"
   ''                clsAdoCon.ExecuteSQL sSQL
   ''            End If
   '            gl.AddDebitCredit cFREIGHT, 0, sCOSjFrtAcct, "", 0, 0, "", sCust
   '
   '            ' Tax
   ''            If cTax <> 0 Then
   ''                iRef = iRef + 1
   ''                sSql = "INSERT INTO JritTable (DCHEAD,DCTRAN,DCREF," _
   ''                    & "DCDEBIT,DCACCTNO,DCDATE,DCCUST,DCINVNO) VALUES('" _
   ''                    & Trim(sJournalID) & "'," _
   ''                    & iTrans & "," _
   ''                    & iRef & "," _
   ''                    & Abs(cTax) & ",'" _
   ''                    & sCOSjTaxAcct & "','" _
   ''                    & Format(txtDte, "mm/dd/yy") & "','" _
   ''                    & sCust & "'," _
   ''                    & lNewInv & ")"
   ''                clsAdoCon.ExecuteSQL sSQL
   ''            End If
   '            gl.AddDebitCredit cTax, 0, sCOSjTaxAcct, "", 0, 0, "", sCust
   '
   '            ' Fed Tax
   ''            If cTax <> 0 Then
   ''                iRef = iRef + 1
   ''                sSql = "INSERT INTO JritTable (DCHEAD,DCTRAN,DCREF," _
   ''                    & "DCDEBIT,DCACCTNO,DCDATE,DCCUST,DCINVNO) VALUES('" _
   ''                    & Trim(sJournalID) & "'," _
   ''                    & iTrans & "," _
   ''                    & iRef & "," _
   ''                    & Abs(cFedTax) & ",'" _
   ''                    & sCOSjFedTaxAcct & "','" _
   ''                    & Format(txtDte, "mm/dd/yy") & "','" _
   ''                    & sCust & "'," _
   ''                    & lNewInv & ")"
   ''                clsAdoCon.ExecuteSQL sSQL
   ''            End If
   '            gl.AddDebitCredit cFedTax, 0, sCOSjFedTaxAcct, "", 0, 0, "", sCust
   '
   '        ' Debit memo routine...
   '        Else
   
   ' Debit A/R
   '            iRef = iRef + 1
   '            sSql = "INSERT INTO JritTable (DCHEAD,DCTRAN,DCREF," _
   '                & "DCDEBIT,DCACCTNO,DCDATE,DCCUST,DCINVNO) VALUES('" _
   '                & Trim(sJournalID) & "'," _
   '                & iTrans & "," _
   '                & iRef & "," _
   '                & (cTotal + cTax + cFREIGHT) & ",'" _
   '                & sCOSjARAcct & "','" _
   '                & Format(txtDte, "mm/dd/yy") & "','" _
   '                & sCust & "'," _
   '                & lNewInv & ")"
   '            clsAdoCon.ExecuteSQL sSQL
   
   gl.AddDebitCredit cTotal + cTax + cFedTax + cFREIGHT, 0, sCOSjARAcct, "", 0, 0, "", sCust
   
   ' Credit Revenue
   '            iRef = iRef + 1
   '            sSql = "INSERT INTO JritTable (DCHEAD,DCTRAN,DCREF," _
   '                & "DCCREDIT,DCACCTNO,DCDATE,DCCUST,DCINVNO) VALUES('" _
   '                & Trim(sJournalID) & "'," _
   '                & iTrans & "," _
   '                & iRef & "," _
   '                & cTotal & ",'" _
   '                & sCOCrRevAcct & "','" _
   '                & Format(txtDte, "mm/dd/yy") & "','" _
   '                & sCust & "'," _
   '                & lNewInv & ")"
   '            clsAdoCon.ExecuteSQL sSQL
   gl.AddDebitCredit 0, cTotal, sCOCrRevAcct, "", 0, 0, "", sCust
   
   ' Freight
   '            If cFREIGHT <> 0 Then
   '                iRef = iRef + 1
   '                sSql = "INSERT INTO JritTable (DCHEAD,DCTRAN,DCREF," _
   '                    & "DCCREDIT,DCACCTNO,DCDATE,DCCUST,DCINVNO) VALUES('" _
   '                    & Trim(sJournalID) & "'," _
   '                    & iTrans & "," _
   '                    & iRef & "," _
   '                    & cFREIGHT & ",'" _
   '                    & sCOSjFrtAcct & "','" _
   '                    & Format(txtDte, "mm/dd/yy") & "','" _
   '                    & sCust & "'," _
   '                    & lNewInv & ")"
   '                clsAdoCon.ExecuteSQL sSQL
   '            End If
   gl.AddDebitCredit 0, cFREIGHT, sCOSjFrtAcct, "", 0, 0, "", sCust
   
   ' Tax
   '            If cTax <> 0 Then
   '                iRef = iRef + 1
   '                sSql = "INSERT INTO JritTable (DCHEAD,DCTRAN,DCREF," _
   '                    & "DCCREDIT,DCACCTNO,DCDATE,DCCUST,DCINVNO) VALUES('" _
   '                    & Trim(sJournalID) & "'," _
   '                    & iTrans & "," _
   '                    & iRef & "," _
   '                    & cTax & ",'" _
   '                    & sCOSjTaxAcct & "','" _
   '                    & Format(txtDte, "mm/dd/yy") & "','" _
   '                    & sCust & "'," _
   '                    & lNewInv & ")"
   '                clsAdoCon.ExecuteSQL sSQL
   '            End If
   gl.AddDebitCredit 0, cTax, sCOSjTaxAcct, "", 0, 0, "", sCust
   
   ' Fed Tax
   '            If cTax <> 0 Then
   '                iRef = iRef + 1
   '                sSql = "INSERT INTO JritTable (DCHEAD,DCTRAN,DCREF," _
   '                    & "DCCREDIT,DCACCTNO,DCDATE,DCCUST,DCINVNO) VALUES('" _
   '                    & Trim(sJournalID) & "'," _
   '                    & iTrans & "," _
   '                    & iRef & "," _
   '                    & cFedTax & ",'" _
   '                    & sCOSjFedTaxAcct & "','" _
   '                    & Format(txtDte, "mm/dd/yy") & "','" _
   '                    & sCust & "'," _
   '                    & lNewInv & ")"
   '                clsAdoCon.ExecuteSQL sSQL
   '            End If
   gl.AddDebitCredit 0, cFedTax, sCOSjFedTaxAcct, "", 0, 0, "", sCust
   
   '        End If
   'End If
   
   '    ' No errors, commit transaction to SQL database
   '    If Err = 0 Then
   '        clsADOCon.CommitTrans
   '        sMsg = "Invoice " & cmbPre & Format(lNewInv, "000000") _
   '            & " Was Successfully Posted." & vbCrLf & vbCrLf _
   '            & "Do You Wish To Print It?"
   '        bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   '        If bResponse = vbYes Then
   '            diaARp01a.Visible = False
   '            Show diaARp01a
   '            diaARp01a.cmbInv = lNewInv
   '            sTemp = diaARp01a.lblPrinter
   '            diaARp01a.lblPrinter = lblPrinter
   '            DoEvents
   '            diaARp01a.optPrn = True
   '            diaARp01a.lblPrinter = sTemp
   '            Unload diaARp01a
   '            Me.SetFocus
   '        End If
   '        CloseBoxes
   '        AddInvoice
   '    Else
   '        clsADOCon.RollBackTrans
   '        MsgBox "Could Not Complete The Transaction.", _
   '            vbExclamation, Caption
   '    End If
   
   'test that debits and credits balance.  if they do, commit transaction to db
   Dim success As Boolean
   success = gl.Commit
   
   If success And clsADOCon.ADOErrNum = 0 Then
      clsADOCon.CommitTrans
      sMsg = "Invoice " & cmbPre & Format(lNewInv, "000000") _
             & " Was Successfully Posted." & vbCrLf & vbCrLf _
             & "Do You Wish To Print It?"
      bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
      If bResponse = vbYes Then
         diaARp01a.Visible = False
         Show diaARp01a
         'diaARp01a.cmbInv = lNewInv
         sTemp = diaARp01a.lblPrinter
         diaARp01a.lblPrinter = lblPrinter
         DoEvents                      'forces form_activate, which sets cmbInv to the wrong #
         diaARp01a.cmbInv = lNewInv    'must happen after form activate, which sets cmbinv
         diaARp01a.optPrn = True
         diaARp01a.lblPrinter = sTemp
         Unload diaARp01a
         
'         diaARp01a.Visible = False
'         DoEvents
'         diaARp01a.PrintMemo lNewInv, sType
'         Unload diaARp01a
         
         Me.SetFocus
      End If
      CloseBoxes
      AddInvoice
   Else
      clsADOCon.RollbackTrans
      clsADOCon.ADOErrNum = 0
      MsgBox "Could Not Complete The Transaction.", _
         vbExclamation, Caption
   End If
   
   
   
   ' Reset settings and exit sub
   On Error Resume Next
   cmbTyp.SetFocus
   Exit Sub
   
DiaErr1:
   sProcName = "updateinvoice"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub GetAccount(sControl As String, sCmbAcct As String)
   Dim RdoGlm As ADODB.Recordset
   On Error GoTo DiaErr1
   sControl = UCase(sControl)
   sCmbAcct = Compress(sCmbAcct)
   If sCmbAcct = "" Then Exit Sub
   sSql = "SELECT GLACCTREF,GLACCTNO,GLDESCR FROM GlacTable " _
          & "WHERE GLACCTREF='" & sCmbAcct & "' AND GLINACTIVE=0"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoGlm, ES_FORWARD)
   If bSqlRows Then
      Select Case sControl
         Case "TXTFRA"
            txtFra = "" & Trim(RdoGlm!GLACCTNO)
            lblFra = "" & Trim(RdoGlm!GLDESCR)
         Case "TXTTXA"
            txtTxa = "" & Trim(RdoGlm!GLACCTNO)
            lblTxa = "" & Trim(RdoGlm!GLDESCR)
         Case "CBOFEDTAX"
            cboFedTax = "" & Trim(RdoGlm!GLACCTNO)
            lblFedTax = "" & Trim(RdoGlm!GLDESCR)
         Case Else
            txtReva = "" & Trim(RdoGlm!GLACCTNO)
            lblReva = "" & Trim(RdoGlm!GLDESCR)
      End Select
   Else
      If bGlverify Then
         Select Case sControl
            Case "TXTFRA"
               lblFra = "*** Account Wasn't Found Or Inactive ***"
            Case "TXTTXA"
               lblTxa = "*** Account Wasn't Found Or Inactive ***"
            Case "cboFedTax"
               lblFedTax = "*** Account Wasn't Found Or Inactive ***"
            Case Else
               lblReva = "*** Account Wasn't Found Or Inactive ***"
         End Select
      End If
   End If
   Set RdoGlm = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "addinvoice"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub FillAccounts()
   Dim RdoGlm As ADODB.Recordset
   Dim b As Byte
   Dim i As Integer
   On Error GoTo DiaErr1
   i = -1
   sProcName = "fillaccou"
   sSql = "Qry_FillLowAccounts"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoGlm, ES_FORWARD)
   If bSqlRows Then
      With RdoGlm
         Do Until .EOF
            i = i + 1
            AddComboStr txtFra.hWnd, "" & Trim(!GLACCTNO)
            AddComboStr txtTxa.hWnd, "" & Trim(!GLACCTNO)
            AddComboStr cboFedTax.hWnd, "" & Trim(!GLACCTNO)
            AddComboStr txtReva.hWnd, "" & Trim(!GLACCTNO)
            .MoveNext
         Loop
         .Cancel
      End With
   End If
   bGlverify = GetMemoAccounts()
   If txtFra.ListCount > 0 Then
      txtFra = sCOSjFrtAcct
      GetAccount "txtfra", txtFra
      txtTxa = sCOSjTaxAcct
      GetAccount "txtTxa", txtTxa
      cboFedTax = sCOSjFedTaxAcct
      GetAccount "cboFedTax", cboFedTax
      txtReva = sCOCrRevAcct
      GetAccount "txtReva", txtReva
   End If
   Set RdoGlm = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "fillaccounts"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Public Function GetMemoAccounts() As Byte
   Dim rdoCsh As ADODB.Recordset
   Dim i As Integer
   Dim b As Byte
   Err = 0
   sProcName = "getmemoacct"
   sSql = "SELECT COGLVERIFY,COCRREVACCT,COSJTAXACCT,COSJNFRTACCT," _
          & "COSJARACCT, COFEDTAXACCT FROM ComnTable WHERE COREF=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoCsh, ES_FORWARD)

' MM MsgBox ("SQL Query: " & sSql)
   
   If bSqlRows Then
      With rdoCsh
         If !COGLVERIFY = 1 Then
            For i = 0 To 4
               If IsNull(.Fields(i)) Then
                  b = 1
               Else
                  If Trim(.Fields(i)) = "" Then b = 1
               End If
            Next
            
' MM MsgBox ("Inside recordset")
' MM MsgBox ("Variable Value Before: " & sCOSjARAcct)
            
            If sCOSjARAcct = "" Then sCOSjARAcct = "" & Trim(!COSJARACCT)
            If sCOCrRevAcct = "" Then sCOCrRevAcct = "" & Trim(!COCRREVACCT)
            If sCOSjTaxAcct = "" Then sCOSjTaxAcct = "" & Trim(!COSJTAXACCT)
            If sCOSjFrtAcct = "" Then sCOSjFrtAcct = "" & Trim(!COSJNFRTACCT)
            If sCOSjFedTaxAcct = "" Then sCOSjFedTaxAcct = "" & Trim(!COFEDTAXACCT)
            .Cancel
' MM MsgBox ("DB Column Value: " & Trim(!COSJARACCT))
' MM MsgBox ("Variable Value after : " & sCOSjARAcct)
            GetMemoAccounts = 1
         Else
' MM MsgBox ("GO GL Failed!")
            GetMemoAccounts = 0
         End If
         .Cancel
' MM MsgBox ("End of record set!")
      End With
      If GetMemoAccounts = 1 Then
         If b = 1 Then lblJrn.Visible = True
      End If
   End If
   Set rdoCsh = Nothing
   
' MM MsgBox ("Setting Sales JournalAcc: " & sCOSjARAcct)

   
End Function

' Return byref customers tax rate, tax state, and tax code.
' 10/15/03 (nth)

Private Sub GetSalesTax( _
                        sCust As String, _
                        Optional nTaxRate As Single, _
                        Optional sTaxState As String, _
                        Optional sTaxCode As String)
   
   On Error GoTo DiaErr1
   Dim RdoTax As ADODB.Recordset
   
   sSql = "SELECT TAXCODE,TAXSTATE,TAXRATE FROM CustTable INNER JOIN " _
          & "TxcdTable ON CustTable.CUTAXCODE = TxcdTable.TAXCODE " _
          & "WHERE CUREF = '" & Compress(sCust) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoTax)
   
   If bSqlRows Then
      With RdoTax
         nTaxRate = !TAXRATE
         sTaxCode = "" & Trim(!taxCode)
         sTaxState = "" & Trim(!taxState)
         .Cancel
      End With
   End If
   Set RdoTax = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub SaveOptions()
   Dim sOptions As String
   SaveSetting "Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter
End Sub

Private Sub GetOptions()
   On Error Resume Next
   sCOSjFrtAcct = GetSetting("Esi2000", "EsiFina", "MemotxtFra", sCOSjFrtAcct)
   sCOSjTaxAcct = GetSetting("Esi2000", "EsiFina", "MemoTxtTxa", sCOSjTaxAcct)
   sCOCrRevAcct = GetSetting("Esi2000", "EsiFina", "MemoTxtRev", sCOCrRevAcct)
   fedTaxRate = CCur(GetSetting("Esi2000", "EsiFina", "MemoFedTaxRate", fedTaxRate))
   lblPrinter = GetSetting("Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter)
   If lblPrinter = "" Then lblPrinter = "Default Printer"
End Sub

Private Sub cmdComments_Click()
'bbs changes from the Comments form to the SysComments form on 6/28/2010 for Ticket #31511
   'Add one of these to the form and go
   If cmdComments Then
      'The Default is txtCmt and need not be included
      'Use Select Case cmdCopy to add your own
      SysComments.lblControl = "txtCmt"
      'See List For Index
      SysComments.lblListIndex = 4
      SysComments.Show
      cmdComments = False
   End If
End Sub


Private Sub GetNextInvoice()
'   Dim RdoInv As ADODB.RecordSet
'   On Error GoTo DiaErr1
'   sSql = "SELECT MAX(INVNO) FROM CihdTable"
'   bSqlRows = clsAdoCon.GetDataSet(sSql,RdoInv)
'   If bSqlRows Then
'      If IsNull(RdoInv.Fields(0)) Then
'         txtInv = "000001"
'      Else
'         lNewInv = (RdoInv.Fields(0) + 1)
'         txtInv = Format(lNewInv, "000000")
'      End If
'   Else
'      txtInv = "000001"
'   End If
'   Set RdoInv = Nothing
'   Exit Sub
'
'DiaErr1:
'   sProcName = "getnextinv"
'   CurrError.Number = Err.Number
'   CurrError.Description = Err.Description
'   DoModuleErrors Me

   Dim inv As New ClassARInvoice
   lNewInv = inv.GetNextInvoiceNumber
   txtInv = Format(lNewInv, "000000")

End Sub

Private Function IsTaxCalculatedPerLineItem() As Boolean
   Dim rdo As ADODB.Recordset
   Dim result As Boolean
   
   result = False
   
   On Error GoTo done
   sSql = "SELECT TaxPerItem from ComnTable"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdo, ES_FORWARD)
   If bSqlRows Then
      result = rdo.Fields(0)
   End If
   Set rdo = Nothing
done:
   IsTaxCalculatedPerLineItem = result
End Function
