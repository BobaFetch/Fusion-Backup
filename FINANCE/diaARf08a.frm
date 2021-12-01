VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form diaARf08a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Export AR Invoice Activity To QuickBooks ® IIF"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5640
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtTaxItem 
      Height          =   285
      Left            =   3000
      TabIndex        =   7
      Top             =   4200
      Width           =   2535
   End
   Begin VB.TextBox txtFrtAcct 
      Height          =   285
      Left            =   3000
      TabIndex        =   8
      Top             =   4680
      Width           =   2535
   End
   Begin VB.TextBox txtFrtItm 
      Height          =   285
      Left            =   3000
      TabIndex        =   9
      Top             =   5040
      Width           =   2535
   End
   Begin VB.CheckBox optVew 
      Caption         =   "vew"
      Height          =   255
      Left            =   3000
      TabIndex        =   32
      Top             =   1200
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdVew 
      Height          =   320
      Left            =   2520
      Picture         =   "diaARf08a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   31
      TabStop         =   0   'False
      ToolTipText     =   "Show Invoices Listed For Export"
      Top             =   1200
      Width           =   350
   End
   Begin VB.TextBox txtTaxVend 
      Height          =   285
      Left            =   3000
      TabIndex        =   6
      Top             =   3840
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   0
      TabIndex        =   25
      Top             =   2040
      Width           =   5655
   End
   Begin VB.TextBox txtQBDis 
      Height          =   285
      Left            =   3000
      TabIndex        =   4
      Top             =   3000
      Width           =   2535
   End
   Begin VB.TextBox txtQBTax 
      Height          =   285
      Left            =   3000
      TabIndex        =   5
      Top             =   3480
      Width           =   2535
   End
   Begin VB.TextBox txtQBSales 
      Height          =   285
      Left            =   3000
      TabIndex        =   3
      Top             =   2640
      Width           =   2535
   End
   Begin VB.TextBox txtQBAR 
      Height          =   285
      Left            =   3000
      TabIndex        =   2
      Top             =   2280
      Width           =   2535
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5160
      Top             =   720
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   6375
      FormDesignWidth =   5640
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5160
      TabIndex        =   12
      TabStop         =   0   'False
      ToolTipText     =   "Browse For File Location"
      Top             =   5520
      Width           =   375
   End
   Begin VB.TextBox txtPath 
      Height          =   285
      Left            =   1320
      TabIndex        =   10
      Top             =   5520
      Width           =   3735
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4680
      TabIndex        =   11
      ToolTipText     =   "Build QuickBooks Export"
      Top             =   6000
      Width           =   875
   End
   Begin VB.ComboBox txtEnd 
      Height          =   315
      Left            =   1800
      TabIndex        =   1
      Tag             =   "4"
      Top             =   720
      Width           =   1095
   End
   Begin VB.ComboBox txtStart 
      Height          =   315
      Left            =   1800
      TabIndex        =   0
      Tag             =   "4"
      Top             =   360
      Width           =   1095
   End
   Begin ComctlLib.ProgressBar prg1 
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   6000
      Visible         =   0   'False
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Close"
      Height          =   435
      Left            =   4680
      TabIndex        =   13
      TabStop         =   0   'False
      ToolTipText     =   "Save And Exit"
      Top             =   120
      Width           =   875
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   18
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
      PictureUp       =   "diaARf08a.frx":04DA
      PictureDn       =   "diaARf08a.frx":0620
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Item"
      Height          =   285
      Index           =   13
      Left            =   120
      TabIndex        =   35
      Top             =   4200
      Width           =   1905
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Freight Account"
      Height          =   285
      Index           =   12
      Left            =   120
      TabIndex        =   34
      Top             =   4680
      Width           =   1905
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Freight Item"
      Height          =   285
      Index           =   8
      Left            =   120
      TabIndex        =   33
      Top             =   5040
      Width           =   1905
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Payable Vendor"
      Height          =   285
      Index           =   11
      Left            =   120
      TabIndex        =   30
      Top             =   3840
      Width           =   1905
   End
   Begin VB.Label lblDte 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   4560
      TabIndex        =   29
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label lblLast 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1800
      TabIndex        =   28
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Last Export Date"
      Height          =   285
      Index           =   10
      Left            =   3120
      TabIndex        =   27
      Top             =   1680
      Width           =   1305
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Last Invoice Exported"
      Height          =   285
      Index           =   9
      Left            =   120
      TabIndex        =   26
      Top             =   1680
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Discount Account"
      Height          =   285
      Index           =   7
      Left            =   120
      TabIndex        =   24
      Top             =   3000
      Width           =   1305
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Account"
      Height          =   285
      Index           =   6
      Left            =   120
      TabIndex        =   23
      Top             =   3480
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Account"
      Height          =   285
      Index           =   5
      Left            =   120
      TabIndex        =   22
      Top             =   2640
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "AR Account"
      Height          =   285
      Index           =   4
      Left            =   120
      TabIndex        =   21
      Top             =   2280
      Width           =   1305
   End
   Begin VB.Label lblFound 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1800
      TabIndex        =   20
      Tag             =   "1"
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Invoices Found"
      Height          =   285
      Index           =   3
      Left            =   120
      TabIndex        =   19
      Top             =   1200
      Width           =   1305
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Disk Path\File"
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   17
      Top             =   5520
      Width           =   1185
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "End Date"
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   15
      Top             =   720
      Width           =   825
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Start Date"
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   14
      Top             =   360
      Width           =   825
   End
End
Attribute VB_Name = "diaARf08a"
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
' diaARf08a - Export AR Activity To QuickBooks
'
' Notes:
'
' Created: 6/17/02 (nth)
' Revisions:
'   06/18/02 (nth) Added Invoices Found
'   06/25/02 (nth) Added SaveOptions to remember QuickBooks AR accounts
'   06/25/02 (nth) Added browse button functionality
'   06/26/02 (nth) Changed QB Account storage to DB rather than registry
'   06/26/02 (nth) Allow multiple QB companies (not yet turned on)
'   06/26/02 (nth) Added invoice list popup
'   06/27/02 (nth) Added freight invoice line item
'   07/09/02 (nth) Fixed offset error with account and customer
'   07/09/02 (nth) Added QBTAXITEM
'   08/20/02 (nth) Fixed error with exporting CM,DM memos
'   08/27/02 (nth) Fixed error with terms not exporting with PS invoice
'   08/27/02 (nth) Added Invoice Due Date
'   08/28/02 (nth) Per requiriments defined by Quickbooks terms used on PS invoice are from customer
'   12/30/02 (nth) Fixed rutime cursor error caused by empty QkbkTable.
'   03/17/05 cjs Fixed QkbkTable name in IniQb
'*************************************************************************************

Dim bOnLoad As Byte
Dim sInvHdr(22) As String
Dim sInvItHdr(19) As String
Dim rdoQB As ADODB.Recordset
Dim bGoodQB As Byte

Const sQBEndTrans = "!ENDTRNS"

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'*************************************************************************************

Private Sub cmdBrowse_Click()
   MdiSect.Cdi.InitDir = txtPath
   MdiSect.Cdi.ShowSave
   txtPath = MdiSect.Cdi.FileName
End Sub

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdGo_Click()
   BuildQBExport
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      SelectHelpTopic Me, "No Subject Help"
      cmdHlp = False
      MouseCursor 0
   End If
End Sub

Private Sub cmdVew_Click()
   Dim RdoInv As ADODB.Recordset
   Dim sItem As String
   
   VewInv.lblCaption = "Invoices Selected For Export"
   sSql = "SELECT INVNO,INVDATE,CUNICKNAME,CUQBNAME,INVTOTAL " _
          & "FROM CihdTable INNER JOIN CustTable ON CihdTable.INVCUST = CustTable.CUREF " _
          & "WHERE INVDATE >= '" & txtstart & "' AND INVDATE <='" & txtEnd & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoInv)
   If bSqlRows Then
      With RdoInv
         While Not .EOF
            sItem = "" & Trim(!invno) & vbTab & Space(5) _
                    & "" & Format(Trim(!INVDATE), "mm/dd/yy") & vbTab _
                    & "" & Trim(!CUNICKNAME) & vbTab & vbTab _
                    & Format("" & Trim(!INVTOTAL), "#,###,##0.00")
            VewInv.lstInv.AddItem sItem
            .MoveNext
         Wend
      End With
   End If
   
   Set RdoInv = Nothing
   VewInv.Show
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad = 1 Then
      bOnLoad = 0
      IniQB
      GetInvoices
   End If
   
   If optVew.Value = vbChecked Then
      Unload VewInv
      optVew.Value = vbUnchecked
   End If
   
   MouseCursor 0
End Sub

Private Sub Form_Load()
   Dim rdoStr As ADODB.Recordset
   
   'FormLoad Me, ES_DONTLIST, ES_DONTLIST
   FormLoad Me, ES_DONTLIST
   FormatControls
   bOnLoad = 1
   
   ' Default starting date is one day after most recent GL period closed
   sSql = "SELECT MAX(GJPOST) FROM GjhdTable WHERE GJPOSTED = 1"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoStr)
   
   If Not IsNull(rdoStr.Fields(0)) Then
      txtstart = Format(DateAdd("d", 1, "" _
                 & Trim(rdoStr.Fields(0))), "mm/dd/yy")
   Else
      txtstart = Format(Now, "mm/01/yy")
   End If
   txtEnd = Format(Now, "mm/dd/yy")
   Set rdoStr = Nothing
   
   'QuickBooks !TRNS
   sInvHdr(0) = "!TRNS"
   sInvHdr(1) = "TRNSID"
   sInvHdr(2) = "TRNSTYPE"
   sInvHdr(3) = "DATE"
   sInvHdr(4) = "ACCNT"
   sInvHdr(5) = "NAME"
   sInvHdr(6) = "CLASS"
   sInvHdr(7) = "AMOUNT"
   sInvHdr(8) = "DOCNUM"
   sInvHdr(9) = "MEMO"
   sInvHdr(10) = "CLEAR"
   sInvHdr(11) = "TOPRINT"
   sInvHdr(12) = "NAMEISTAXABLE"
   sInvHdr(13) = "ADDR1"
   sInvHdr(14) = "ADDR2"
   sInvHdr(15) = "ADDR3"
   sInvHdr(16) = "ADDR4"
   sInvHdr(17) = "ADDR5"
   sInvHdr(18) = "DUEDATE"
   sInvHdr(19) = "TERMS"
   sInvHdr(20) = "PAID"
   sInvHdr(21) = "PAYMETH"
   sInvHdr(22) = "SHIPDATE"
   
   'QuickBooks !SPL
   sInvItHdr(0) = "!SPL"
   sInvItHdr(1) = "SPLID"
   sInvItHdr(2) = "TRNSTYPE"
   sInvItHdr(3) = "DATE"
   sInvItHdr(4) = "ACCNT"
   sInvItHdr(5) = "NAME"
   sInvItHdr(6) = "CLASS"
   sInvItHdr(7) = "AMOUNT"
   sInvItHdr(8) = "DOCNUM"
   sInvItHdr(9) = "MEMO"
   sInvItHdr(10) = "CLEAR"
   sInvItHdr(11) = "QNTY"
   sInvItHdr(12) = "PRICE"
   sInvItHdr(13) = "INVITEM"
   sInvItHdr(14) = "PAYMETH"
   sInvItHdr(15) = "TAXABLE"
   sInvItHdr(16) = "VALADJ"
   sInvItHdr(17) = "REIMBEXP"
   sInvItHdr(18) = "SERVICEDATE"
   sInvItHdr(19) = "EXTRA"
   
   MdiSect.Cdi.Filter = "Intuit Interchange Format  *.IIF"
   MdiSect.Cdi.FilterIndex = 1
   MdiSect.Cdi.DefaultExt = "IIF"
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   FormUnload
   Set rdoQB = Nothing
   Set diaARf08a = Nothing
End Sub

Public Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Private Sub txtend_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtEnd_LostFocus()
   txtEnd = CheckDate(txtEnd)
   GetInvoices
End Sub

Private Sub txtFrtItm_LostFocus()
   txtFrtItm = CheckLen(txtFrtItm, 50)
   If bGoodQB Then
      On Error Resume Next
      rdoQB!QBARFRT = Trim(txtFrtItm)
      rdoQB.Update
      If Err > 0 Then ValidateEdit Me
   End If
End Sub

Private Sub txtFrtAcct_LostFocus()
   txtFrtAcct = CheckLen(txtFrtAcct, 50)
   If bGoodQB Then
      On Error Resume Next
      rdoQB!QBARFRTACCT = Trim(txtFrtAcct)
      rdoQB.Update
      If Err > 0 Then ValidateEdit Me
   End If
End Sub

Private Sub txtPath_LostFocus()
   txtPath = CheckLen(txtPath, 256)
   If bGoodQB Then
      On Error Resume Next
      rdoQB!QBEXPPATH = Trim(txtPath)
      rdoQB.Update
      If Err > 0 Then ValidateEdit Me
   End If
End Sub

Private Sub txtQBAR_LostFocus()
   txtQBAR = CheckLen(txtQBAR, 50)
   If bGoodQB Then
      On Error Resume Next

      rdoQB!QBARACCT = Trim(txtQBAR)
      rdoQB.Update
      If Err > 0 Then ValidateEdit Me
   End If
End Sub

Private Sub txtQBDis_LostFocus()
   txtQBDis = CheckLen(txtQBDis, 50)
   If bGoodQB Then
      On Error Resume Next
      rdoQB!QBDISACCT = Trim(txtQBDis)
      rdoQB.Update
      If Err > 0 Then ValidateEdit Me
   End If
End Sub

Private Sub txtQBSales_LostFocus()
   txtQBSales = CheckLen(txtQBSales, 50)
   If bGoodQB Then
      On Error Resume Next
      rdoQB!QBSALESACCT = Trim(txtQBSales)
      rdoQB.Update
      If Err > 0 Then ValidateEdit Me
   End If
End Sub

Private Sub txtQBTax_LostFocus()
   txtQBTax = CheckLen(txtQBTax, 50)
   If bGoodQB Then
      On Error Resume Next
      rdoQB!QBTAXACCT = Trim(txtQBTax)
      rdoQB.Update
      If Err > 0 Then ValidateEdit Me
   End If
End Sub

Private Sub txtstart_DropDown()
   ShowCalendar Me
End Sub

Private Sub GetInvoices()
   Dim RdoInv As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT COUNT(INVNO) FROM CihdTable WHERE INVDATE >= '" & txtstart _
          & "' AND INVDATE <= '" & txtEnd & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoInv, ES_FORWARD)
   If RdoInv.Fields(0) > 0 Then
      lblFound = RdoInv.Fields(0)
      cmdGo.enabled = True
   Else
      lblFound = 0
      cmdGo.enabled = False
   End If
   Set RdoInv = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getinvoices"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub BuildQBExport()
   Dim iFile As Integer
   Dim sFileName As String
   Dim RdoInv As ADODB.Recordset 'Invoice Header
   Dim RdoIt As ADODB.Recordset 'Invoice Items
   Dim RdoTerms As ADODB.Recordset 'Invoice Terms
   Dim AdoQry1 As ADODB.Command 'SO
   Dim AdoParameter1 As ADODB.Parameter
  
   Dim AdoQry2 As ADODB.Command 'PS
   Dim AdoParameter2 As ADODB.Parameter
   
   
   Dim i As Integer
   Dim sCust As String
   Dim sAddr(4) As String
   Dim sPIF As String * 1
   Dim sTaxable As String * 1
   Dim sMsg As String
   Dim nQty As Single
   Dim nPrice As Single
   Dim sPart As String
   Dim sbuf As String
   
   Dim iLastInv As Long
   Dim sLastInvPre As String
   
   Dim bNoItems As Byte
   Dim sTerms As String
   Dim sDue As String
   
   On Error GoTo DiaErr1
   MouseCursor 13
   iFile = FreeFile
   sFileName = Trim(txtPath)
   
   'SO
   sSql = "SELECT PARTNUM AS Part,PADESC As Desciption, ITQTY AS Qty,ITDOLLARS AS Price,ITCOMMENTS,ITNUMBER as Item " _
          & "FROM SoitTable INNER JOIN PartTable ON SoitTable.ITPART = PartTable.PARTREF " _
          & "INNER JOIN CihdTable ON SoitTable.ITINV = CihdTable.INVNO " _
          & "WHERE INVNO = ? "
   Set AdoQry1 = New ADODB.Command
   AdoQry1.CommandText = sSql
   
   Set AdoParameter1 = New ADODB.Parameter
   AdoParameter1.Type = adChar
   AdoParameter1.SIZE = 8
   AdoQry1.parameters.Append AdoParameter1
   
   
   'PS
   sSql = "SELECT PIQTY AS Qty,PIITNO As Item,PARTNUM AS Part,PADESC As Desciption,PICOMMENTS,PISELLPRICE AS Price " _
          & "FROM PartTable INNER JOIN PsitTable ON PartTable.PARTREF = PsitTable.PIPART INNER JOIN " _
          & "CihdTable ON PsitTable.PIPACKSLIP = CihdTable.INVPACKSLIP " _
          & "WHERE INVNO = ? "
   
   Set AdoQry2 = New ADODB.Command
   AdoQry2.CommandText = sSql
   
   Set AdoParameter2 = New ADODB.Parameter
   AdoParameter2.Type = adChar
   AdoParameter2.SIZE = 8
   AdoQry2.parameters.Append AdoParameter2
   
   'Invoice Headers
   sSql = "SELECT INVNO,INVPRE,INVDATE,INVSHIPDATE,INVSTADR,INVARACCT,INVTAX,INVFREIGHT,CUQBNAME," _
          & "INVTAXACCT,INVTOTAL,INVPIF,INVTYPE,INVCOMMENTS,INVSO,INVCUST " _
          & "FROM CihdTable " _
          & "INNER JOIN CustTable ON CihdTable.INVCUST = CustTable.CUREF " _
          & "WHERE INVDATE >= '" & txtstart & "' AND INVDATE <='" & txtEnd & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoInv, ES_STATIC)
   
   If bSqlRows Then
      prg1.Visible = True
      prg1.max = Val(lblFound)
      Open sFileName For Output As iFile
      
      With RdoInv
         sbuf = ""
         For i = 0 To UBound(sInvHdr)
            sbuf = sbuf & sInvHdr(i) & vbTab
         Next
         Print #iFile, sbuf
         
         sbuf = ""
         For i = 0 To UBound(sInvItHdr)
            sbuf = sbuf & sInvItHdr(i) & vbTab
         Next
         Print #iFile, sbuf
         Print #iFile, sQBEndTrans
         
         Dim ct As Integer
         
         While Not .EOF
         
            ct = ct + 1
            Debug.Print ct
            
            If !INVPIF = 0 Then sPIF = "N" Else sPIF = "Y"
            If !INVTAX = 0 Then sTaxable = "N" Else sTaxable = "Y"
            
            sTerms = ""
            sDue = ""
            
            Select Case !INVTYPE
               Case "SO"
                  ' Invoiced From Sales Order
                  AdoQry1.parameters(0).Value = !invno
                  bSqlRows = clsADOCon.GetQuerySet(RdoIt, AdoQry1)
                  sSql = "SELECT INVDAYS,INVARDISC,INVNETDAYS FROM CihdTable WHERE INVNO = " & !invno
                  bSqlRows = clsADOCon.GetDataSet(sSql, RdoTerms)
                  If bSqlRows Then
                     sTerms = "" & RdoTerms!INVARDISC & "% " & RdoTerms!INVDAYS & " Net " & RdoTerms!INVNETDAYS
                     sDue = Format(DateAdd("d", Val("" & RdoTerms!INVNETDAYS), !INVDATE), "mm/dd/yyyy")
                  End If
                  bNoItems = False
               Case "PS"
                  ' Invoiced From Packslip
                  AdoQry2.parameters(0).Value = !invno
                  bSqlRows = clsADOCon.GetQuerySet(RdoIt, AdoQry2)
                  
                  sSql = "SELECT CUDAYS,CUARDISC,CUNETDAYS FROM CustTable WHERE CUREF = '" & Trim(!INVCUST) & "'"
                  bSqlRows = clsADOCon.GetDataSet(sSql, RdoTerms)
                  If bSqlRows Then
                     sTerms = "" & RdoTerms!CUARDISC & "% " & RdoTerms!CUDAYS & " Net " & RdoTerms!CUNETDAYS
                     sDue = Format(DateAdd("d", Val("" & RdoTerms!CUNETDAYS), !INVDATE), "mm/dd/yyyy")
                  End If
                  bNoItems = False
               Case Else
                  ' Credit/Debit Memo
                  bNoItems = True
            End Select
            Set RdoTerms = Nothing
            
            'Output invoice header
            ' * Not outputed from ES/2002
            sbuf = Right(sInvHdr(0), Len(sInvHdr(0)) - 1) & vbTab
            sbuf = sbuf & "" & vbTab
            sbuf = sbuf & "INVOICE" & vbTab
            sbuf = sbuf & "" & Format(Trim(!INVDATE), "m/d/yyyy") & vbTab
            sbuf = sbuf & "" & Trim(txtQBAR) & vbTab
            sbuf = sbuf & "" & Trim(!CUQBNAME) & vbTab '
            sbuf = sbuf & "" & vbTab
            sbuf = sbuf & "" & Trim(!INVTOTAL) & vbTab
            sbuf = sbuf & "" & !invno & vbTab
            sbuf = sbuf & "" & Trim(!INVCOMMENTS) & vbTab
            sbuf = sbuf & "" & vbTab
            sbuf = sbuf & "" & vbTab
            sbuf = sbuf & "" & sTaxable & vbTab
            sbuf = sbuf & "" & Trim(sAddr(0)) & vbTab
            sbuf = sbuf & "" & Trim(sAddr(1)) & vbTab
            sbuf = sbuf & "" & Trim(sAddr(2)) & vbTab
            sbuf = sbuf & "" & Trim(sAddr(3)) & vbTab
            sbuf = sbuf & "" & Trim(sAddr(4)) & vbTab
            sbuf = sbuf & "" & sDue & vbTab
            sbuf = sbuf & "" & sTerms & vbTab
            sbuf = sbuf & "" & sPIF & vbTab
            sbuf = sbuf & "" & vbTab
            sbuf = sbuf & "" & Format(!INVSHIPDATE, "m/d/yyyy") & vbTab
            Print #iFile, sbuf
            
            ' Output invoice items
            ' * Not outputed from ES/2002
            If bNoItems = False Then
               While Not RdoIt.EOF
                  sbuf = Right(sInvItHdr(0), Len(sInvItHdr(0)) - 1) & vbTab 'ID
                  sbuf = sbuf & "" & vbTab 'SPLID
                  sbuf = sbuf & "INVOICE" & vbTab 'Type
                  sbuf = sbuf & "" & Trim(!INVDATE) & vbTab 'Date
                  sbuf = sbuf & "" & txtQBSales & vbTab 'Accnt
                  sbuf = sbuf & "" & !CUQBNAME & vbTab 'Customer
                  sbuf = sbuf & "" & vbTab 'Class *
                  sbuf = sbuf & "" & (RdoIt!price * RdoIt!qty) * -1 & vbTab 'Amount
                  sbuf = sbuf & "" & vbTab 'Docnum *
                  sbuf = sbuf & "" & Trim(RdoIt!Desciption) & vbTab 'Memo *
                  sbuf = sbuf & "" & vbTab 'Clear *
                  sbuf = sbuf & "" & (RdoIt!qty) * -1 & vbTab 'Qty
                  sbuf = sbuf & "" & RdoIt!price & vbTab 'Price
                  sbuf = sbuf & "" & RdoIt!part & vbTab 'Inv Item
                  sbuf = sbuf & "" & vbTab
                  sbuf = sbuf & "N" & vbTab 'Taxable
                  sbuf = sbuf & "N" & vbTab 'VALADJ *
                  sbuf = sbuf & "NOTHING" & vbTab 'REIMBEXP
                  sbuf = sbuf & "" & vbTab 'SERVICEDATE *
                  sbuf = sbuf & "" & vbTab 'EXTRA *
                  Print #iFile, sbuf
                  RdoIt.MoveNext
               Wend
            End If
            
            If !INVFREIGHT <> 0 Then
               sbuf = Right(sInvItHdr(0), Len(sInvItHdr(0)) - 1) & vbTab
               sbuf = sbuf & vbTab
               sbuf = sbuf & "INVOICE" & vbTab
               sbuf = sbuf & "" & Trim(!INVDATE) & vbTab
               sbuf = sbuf & "" & txtFrtAcct & vbTab
               sbuf = sbuf & "" & Trim(!CUQBNAME) & vbTab
               sbuf = sbuf & "" & vbTab
               sbuf = sbuf & "" & (!INVFREIGHT * -1) & vbTab
               sbuf = sbuf & "" & vbTab
               sbuf = sbuf & "" & "Invoice Freight" & vbTab
               sbuf = sbuf & "N" & vbTab
               sbuf = sbuf & "" & vbTab
               sbuf = sbuf & "" & !INVFREIGHT & vbTab
               sbuf = sbuf & "" & Trim(txtFrtItm) & vbTab
               sbuf = sbuf & "" & vbTab
               sbuf = sbuf & "N" & vbTab
               sbuf = sbuf & "N" & vbTab
               sbuf = sbuf & "NOTHING" & vbTab
               sbuf = sbuf & "" & vbTab
               sbuf = sbuf & "" & vbTab
               Print #iFile, sbuf
            End If
            
            sbuf = Right(sInvItHdr(0), Len(sInvItHdr(0)) - 1) & vbTab
            sbuf = sbuf & vbTab
            sbuf = sbuf & "INVOICE" & vbTab
            sbuf = sbuf & "" & Trim(!INVDATE) & vbTab
            sbuf = sbuf & "" & txtQBTax & vbTab
            sbuf = sbuf & "" & txtTaxVend & vbTab
            sbuf = sbuf & "" & vbTab
            sbuf = sbuf & "" & (!INVTAX * -1) & vbTab
            sbuf = sbuf & "" & vbTab
            sbuf = sbuf & "" & vbTab
            sbuf = sbuf & "N" & vbTab
            sbuf = sbuf & "" & vbTab
            sbuf = sbuf & "" & vbTab
            sbuf = sbuf & "" & txtTaxItem & vbTab
            sbuf = sbuf & "" & vbTab
            sbuf = sbuf & "N" & vbTab
            sbuf = sbuf & "N" & vbTab
            sbuf = sbuf & "NOTHING" & vbTab
            sbuf = sbuf & "" & vbTab
            sbuf = sbuf & "AUTOSTAX" & vbTab
            Print #iFile, sbuf
            
            Print #iFile, Right(sQBEndTrans, Len(sQBEndTrans) - 1)
            Set RdoIt = Nothing
            prg1.Value = prg1.Value + 1
            iLastInv = !invno
            sLastInvPre = !INVPRE
            .MoveNext
         Wend
      End With
      
      ' Remeber last invoice and last date
      rdoQB!QBLASTARINV = iLastInv
      rdoQB!QBLASTAREXP = Format(Now, "mm/dd/yyyy")
      rdoQB.Update
      lblDte = Format("" & Trim(rdoQB!QBLASTAREXP), "m/d/yyyy")
      lblLast = sLastInvPre & Format(iLastInv, "000000")
      
      
      sMsg = "Successfully Exported AR Activity."
      SysMsg sMsg, True
   Else
      sMsg = "Could Not Build QuickBooks" & vbCrLf & "AR Export."
      MsgBox sMsg, vbExclamation, Caption
   End If
   
CleanUp:
   prg1.Value = 0
   prg1.Visible = False
   Close iFile
   Set RdoInv = Nothing
   Set RdoIt = Nothing
   Set AdoParameter1 = Nothing
   Set AdoParameter2 = Nothing
   Set AdoQry1 = Nothing
   Set AdoQry2 = Nothing
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "BuildQBExport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   GoTo CleanUp
End Sub

Private Sub txtstart_LostFocus()
   txtstart = CheckDate(txtstart)
   GetInvoices
End Sub

Public Function GetQBSettings() As Byte
   Dim rdoPre As ADODB.Recordset
   
   On Error GoTo DiaErr1
   sSql = "SELECT * FROM QkbkTable"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoQB, ES_KEYSET)
   If bSqlRows Then
      With rdoQB
         txtQBAR = "" & Trim(!QBARACCT)
         txtQBTax = "" & Trim(!QBTAXACCT)
         txtQBDis = "" & Trim(!QBDISACCT)
         txtQBSales = "" & Trim(!QBSALESACCT)
         txtPath = "" & Trim(!QBEXPPATH)
         txtTaxVend = "" & Trim(!QBTAXVEND)
         txtFrtItm = "" & Trim(!QBARFRT)
         txtFrtAcct = "" & Trim(!QBARFRTACCT)
         txtTaxItem = "" & Trim(!QBTAXITEM)
         lblDte = Format("" & Trim(!QBLASTAREXP), "m/d/yyyy")
         
         If Not IsNull(!QBLASTARINV) Then
            sSql = "SELECT INVPRE FROM CihdTable WHERE INVNO = " & !QBLASTARINV
            bSqlRows = clsADOCon.GetDataSet(sSql, rdoPre)
            If bSqlRows Then
               lblLast = rdoPre!INVPRE & Format("" & Trim(!QBLASTARINV), "000000")
            End If
            Set rdoPre = Nothing
         Else
            lblLast = ""
         End If
      End With
      GetQBSettings = 1
   Else
      GetQBSettings = 0
   End If
   Exit Function
   
DiaErr1:
   sProcName = "getqbsettings"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Public Sub IniQB()
   On Error GoTo DiaErr1
   
   bGoodQB = GetQBSettings
   
   ' first time
   If bGoodQB < 1 Then
      sSql = "INSERT INTO QkbkTable(QBREF) VALUES(1)"
      clsADOCon.ExecuteSQL sSql
      bGoodQB = GetQBSettings
   End If
   
   Exit Sub
   
DiaErr1:
   sProcName = "IniQB"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub txtTaxItem_LostFocus()
   txtTaxItem = CheckLen(txtTaxItem, 50)
   If bGoodQB Then
      On Error Resume Next
      rdoQB!QBTAXITEM = Trim(txtTaxItem)
      rdoQB.Update
      If Err > 0 Then ValidateEdit Me
   End If
End Sub

Private Sub txtTaxVend_LostFocus()
   txtTaxVend = CheckLen(txtTaxVend, 30)
   If bGoodQB Then
      On Error Resume Next
      rdoQB!QBTAXVEND = Trim(txtTaxVend)
      rdoQB.Update
      If Err > 0 Then ValidateEdit Me
   End If
End Sub
