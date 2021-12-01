VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form diaARf07
   BorderStyle = 3 'Fixed Dialog
   Caption = "Export AR Invoice Activity To QuickBooks ® IIF"
   ClientHeight = 6480
   ClientLeft = 45
   ClientTop = 330
   ClientWidth = 5655
   ClipControls = 0 'False
   ControlBox = 0 'False
   LinkTopic = "Form1"
   MaxButton = 0 'False
   MDIChild = -1 'True
   MinButton = 0 'False
   ScaleHeight = 6480
   ScaleWidth = 5655
   ShowInTaskbar = 0 'False
   Begin VB.TextBox txtTaxItem
      Height = 285
      Left = 3000
      TabIndex = 7
      Top = 4200
      Width = 2535
   End
   Begin VB.TextBox txtFrtAcct
      Height = 285
      Left = 3000
      TabIndex = 8
      Top = 4680
      Width = 2535
   End
   Begin VB.TextBox txtFrtItm
      Height = 285
      Left = 3000
      TabIndex = 9
      Top = 5040
      Width = 2535
   End
   Begin VB.CheckBox optVew
      Caption = "vew"
      Height = 255
      Left = 3000
      TabIndex = 33
      Top = 1200
      Visible = 0 'False
      Width = 735
   End
   Begin VB.CommandButton cmdVew
      Height = 320
      Left = 2520
      Picture = "diaARCQB.frx":0000
      Style = 1 'Graphical
      TabIndex = 32
      TabStop = 0 'False
      ToolTipText = "Show Invoices Listed For Export"
      Top = 1200
      Width = 350
   End
   Begin VB.TextBox txtTaxVend
      Height = 285
      Left = 3000
      TabIndex = 6
      Top = 3840
      Width = 2535
   End
   Begin VB.Frame Frame2
      Height = 30
      Left = 120
      TabIndex = 30
      Top = 5400
      Width = 5415
   End
   Begin VB.Frame Frame1
      Height = 30
      Left = 120
      TabIndex = 25
      Top = 2040
      Width = 5415
   End
   Begin VB.TextBox txtQBDis
      Height = 285
      Left = 3000
      TabIndex = 4
      Top = 3000
      Width = 2535
   End
   Begin VB.TextBox txtQBTax
      Height = 285
      Left = 3000
      TabIndex = 5
      Top = 3480
      Width = 2535
   End
   Begin VB.TextBox txtQBSales
      Height = 285
      Left = 3000
      TabIndex = 3
      Top = 2640
      Width = 2535
   End
   Begin VB.TextBox txtQBAR
      Height = 285
      Left = 3000
      TabIndex = 2
      Top = 2280
      Width = 2535
   End
   Begin ResizeLibCtl.ReSize ReSize1
      Left = 5160
      Top = 720
      _Version = 196615
      _ExtentX = 741
      _ExtentY = 741
      _StockProps = 0
      Enabled = -1 'True
      FormMinWidth = 0
      FormMinHeight = 0
      FormDesignHeight = 6480
      FormDesignWidth = 5655
   End
   Begin VB.CommandButton cmdBrowse
      Caption = "."
      BeginProperty Font
      Name = "MS Sans Serif"
      Size = 8.25
      Charset = 0
      Weight = 700
      Underline = 0 'False
      Italic = 0 'False
      Strikethrough = 0 'False
      EndProperty
      Height = 315
      Left = 5160
      TabIndex = 12
      TabStop = 0 'False
      ToolTipText = "Browse For File Location"
      Top = 5640
      Width = 375
   End
   Begin VB.TextBox txtPath
      Height = 285
      Left = 1320
      TabIndex = 10
      Top = 5640
      Width = 3735
   End
   Begin VB.CommandButton cmdGo
      Caption = "Go"
      BeginProperty Font
      Name = "MS Sans Serif"
      Size = 8.25
      Charset = 0
      Weight = 700
      Underline = 0 'False
      Italic = -1 'True
      Strikethrough = 0 'False
      EndProperty
      Height = 315
      Left = 4680
      TabIndex = 11
      ToolTipText = "Build QuickBooks Export"
      Top = 6120
      Width = 875
   End
   Begin VB.ComboBox txtEnd
      Height = 315
      Left = 1800
      TabIndex = 1
      Tag = "4"
      Top = 720
      Width = 1095
   End
   Begin VB.ComboBox txtStart
      Height = 315
      Left = 1800
      TabIndex = 0
      Tag = "4"
      Top = 360
      Width = 1095
   End
   Begin ComctlLib.ProgressBar prg1
      Height = 255
      Left = 120
      TabIndex = 16
      Top = 6120
      Visible = 0 'False
      Width = 4455
      _ExtentX = 7858
      _ExtentY = 450
      _Version = 327682
      Appearance = 1
   End
   Begin VB.CommandButton cmdCan
      Caption = "Close"
      Height = 435
      Left = 4680
      TabIndex = 13
      TabStop = 0 'False
      ToolTipText = "Save And Exit"
      Top = 120
      Width = 875
   End
   Begin Threed.SSRibbon cmdHlp
      Height = 225
      Left = 0
      TabIndex = 18
      ToolTipText = "Subject Help"
      Top = 0
      Width = 255
      _Version = 65536
      _ExtentX = 450
      _ExtentY = 397
      _StockProps = 65
      BackColor = 12632256
      Autosize = 2
      RoundedCorners = 0 'False
      BevelWidth = 0
      Outline = 0 'False
      PictureUp = "diaARCQB.frx":04DA
      PictureDn = "diaARCQB.frx":0620
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Tax Item"
      Height = 285
      Index = 13
      Left = 120
      TabIndex = 36
      Top = 4200
      Width = 1905
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Freight Account"
      Height = 285
      Index = 12
      Left = 120
      TabIndex = 35
      Top = 4680
      Width = 1905
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Freight Item"
      Height = 285
      Index = 8
      Left = 120
      TabIndex = 34
      Top = 5040
      Width = 1905
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Tax Payable Vendor"
      Height = 285
      Index = 11
      Left = 120
      TabIndex = 31
      Top = 3840
      Width = 1905
   End
   Begin VB.Label lblDte
      BorderStyle = 1 'Fixed Single
      Height = 285
      Left = 4560
      TabIndex = 29
      Top = 1680
      Width = 975
   End
   Begin VB.Label lblLast
      BorderStyle = 1 'Fixed Single
      Height = 285
      Left = 1800
      TabIndex = 28
      Top = 1680
      Width = 1095
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Last Export Date"
      Height = 285
      Index = 10
      Left = 3120
      TabIndex = 27
      Top = 1680
      Width = 1305
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Last Invoice Exported"
      Height = 285
      Index = 9
      Left = 120
      TabIndex = 26
      Top = 1680
      Width = 1785
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Discount Account"
      Height = 285
      Index = 7
      Left = 120
      TabIndex = 24
      Top = 3000
      Width = 1305
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Tax Account"
      Height = 285
      Index = 6
      Left = 120
      TabIndex = 23
      Top = 3480
      Width = 1425
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Sales Account"
      Height = 285
      Index = 5
      Left = 120
      TabIndex = 22
      Top = 2640
      Width = 1425
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "AR Account"
      Height = 285
      Index = 4
      Left = 120
      TabIndex = 21
      Top = 2280
      Width = 1305
   End
   Begin VB.Label lblFound
      Alignment = 1 'Right Justify
      BorderStyle = 1 'Fixed Single
      Height = 285
      Left = 1800
      TabIndex = 20
      Tag = "1"
      Top = 1200
      Width = 615
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Invoices Found"
      Height = 285
      Index = 3
      Left = 120
      TabIndex = 19
      Top = 1200
      Width = 1305
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Disk Path\File"
      Height = 285
      Index = 2
      Left = 120
      TabIndex = 17
      Top = 5640
      Width = 1185
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "End Date"
      Height = 285
      Index = 0
      Left = 120
      TabIndex = 15
      Top = 720
      Width = 825
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Start Date"
      Height = 285
      Index = 1
      Left = 120
      TabIndex = 14
      Top = 360
      Width = 825
   End
End
Attribute VB_Name = "diaARf07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************************************************
' diaARCQB - Export AR Activity To QuickBooks
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
'
'*************************************************************************************
Option Explicit

Dim bOnLoad As Byte
Dim sInvHdr(22) As String
Dim sInvItHdr(19) As String
Dim rdoQB As rdoResultset
Dim bGoodQB As Byte

Const sQBEndTrans = "!ENDTRNS"

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

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
   Dim RdoInv As rdoResultset
   Dim sItem As String
   
   VewInv.lblCaption = "Invoices Selected For Export"
   sSql = "SELECT INVNO,INVDATE,CUNICKNAME,CUQBNAME,INVTOTAL " _
          & "FROM CihdTable INNER JOIN CustTable ON CihdTable.INVCUST = CustTable.CUREF " _
          & "WHERE INVDATE >= '" & txtStart & "' AND INVDATE <='" & txtEnd & "'"
   bSqlRows = GetDataSet(RdoInv)
   If bSqlRows Then
      With RdoInv
         While Not .EOF
            sItem = "" & Trim(!INVNO) & vbTab & Space(5) _
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
   If bOnLoad Then
      bOnLoad = False
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
   Dim rdoStr As rdoResultset
   
   SetDiaPos Me
   FormatControls
   bOnLoad = True
   
   ' Default starting date is one day after most recent GL period closed
   sSql = "SELECT MAX(GJPOST) FROM GjhdTable WHERE GJPOSTED = 1"
   bSqlRows = GetDataSet(rdoStr)
   
   If Not IsNull(rdoStr.rdoColumns(0)) Then
      txtStart = Format(DateAdd("d", 1, "" & Trim(rdoStr.rdoColumns(0))), "mm/dd/yy")
   Else
      txtStart = Format(Now, "mm/01/yy")
   End If
   txtEnd = Format(Now, "mm/dd/yy")
   
   
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

Private Sub txtEnd_DropDown()
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
      rdoQB.Edit
      rdoQB!QBARFRT = Trim(txtFrtItm)
      rdoQB.Update
      If Err > 0 Then ValidateEdit Me
   End If
End Sub

Private Sub txtFrtAcct_LostFocus()
   txtFrtAcct = CheckLen(txtFrtAcct, 50)
   If bGoodQB Then
      On Error Resume Next
      rdoQB.Edit
      rdoQB!QBARFRTACCT = Trim(txtFrtAcct)
      rdoQB.Update
      If Err > 0 Then ValidateEdit Me
   End If
End Sub

Private Sub txtPath_LostFocus()
   txtPath = CheckLen(txtPath, 256)
   If bGoodQB Then
      On Error Resume Next
      rdoQB.Edit
      rdoQB!QBEXPPATH = Trim(txtPath)
      rdoQB.Update
      If Err > 0 Then ValidateEdit Me
   End If
End Sub

Private Sub txtQBAR_LostFocus()
   txtQBAR = CheckLen(txtQBAR, 50)
   If bGoodQB Then
      On Error Resume Next
      rdoQB.Edit
      rdoQB!QBARACCT = Trim(txtQBAR)
      rdoQB.Update
      If Err > 0 Then ValidateEdit Me
   End If
End Sub

Private Sub txtQBDis_LostFocus()
   txtQBDis = CheckLen(txtQBDis, 50)
   If bGoodQB Then
      On Error Resume Next
      rdoQB.Edit
      rdoQB!QBDISACCT = Trim(txtQBDis)
      rdoQB.Update
      If Err > 0 Then ValidateEdit Me
   End If
End Sub

Private Sub txtQBSales_LostFocus()
   txtQBSales = CheckLen(txtQBSales, 50)
   If bGoodQB Then
      On Error Resume Next
      rdoQB.Edit
      rdoQB!QBSALESACCT = Trim(txtQBSales)
      rdoQB.Update
      If Err > 0 Then ValidateEdit Me
   End If
End Sub

Private Sub txtQBTax_LostFocus()
   txtQBTax = CheckLen(txtQBTax, 50)
   If bGoodQB Then
      On Error Resume Next
      rdoQB.Edit
      rdoQB!QBTAXACCT = Trim(txtQBTax)
      rdoQB.Update
      If Err > 0 Then ValidateEdit Me
   End If
End Sub

Private Sub txtstart_DropDown()
   ShowCalendar Me
End Sub

Private Sub GetInvoices()
   Dim RdoInv As rdoResultset
   sSql = "SELECT COUNT(INVNO) FROM CihdTable WHERE INVDATE >= '" & txtStart _
          & "' AND INVDATE <= '" & txtEnd & "'"
   bSqlRows = GetDataSet(RdoInv, ES_FORWARD)
   If RdoInv.rdoColumns(0) > 0 Then
      lblFound = RdoInv.rdoColumns(0)
      cmdGo.Enabled = True
   Else
      lblFound = 0
      cmdGo.Enabled = False
   End If
   Set RdoInv = Nothing
End Sub

Private Sub BuildQBExport()
   Dim iFile As Integer
   Dim sFileName As String
   Dim RdoInv As rdoResultset 'Invoice Header
   Dim RdoIt As rdoResultset 'Invoice Items
   Dim RdoTerms As rdoResultset 'Invoice Terms
   Dim RdoQry1 As rdoQuery 'SO
   Dim RdoQry2 As rdoQuery 'PS
   Dim RdoQry3 As rdoQuery 'DM
   Dim RdoQry4 As rdoQuery 'CR
   
   Dim i As Integer
   Dim sCust As String
   Dim sAddr(4) As String
   Dim sPIF As String * 1
   Dim sTaxable As String * 1
   Dim smsg As String
   Dim nQty As Single
   Dim nPrice As Single
   Dim sPart As String
   Dim sBuf As String
   
   Dim iLastInv As Integer
   Dim sLastInvPre As String
   
   Dim bNoItems As Byte
   Dim sTerms As String
   Dim sDue As String
   
   On Error GoTo DiaErr1
   'On Error GoTo 0
   
   MouseCursor 13
   iFile = FreeFile
   sFileName = Trim(txtPath)
   
   'SO
   sSql = "SELECT PARTNUM AS Part,PADESC As Desciption, ITQTY AS Qty,ITDOLLARS AS Price,ITCOMMENTS,ITNUMBER as Item " _
          & "FROM SoitTable INNER JOIN PartTable ON SoitTable.ITPART = PartTable.PARTREF " _
          & "INNER JOIN CihdTable ON SoitTable.ITINV = CihdTable.INVNO " _
          & "WHERE INVNO = ? "
   Set RdoQry1 = RdoCon.CreateQuery("", sSql)
   
   'PS
   sSql = "SELECT PIQTY AS Qty,PIITNO As Item,PARTNUM AS Part,PADESC As Desciption,PICOMMENTS,PISELLPRICE AS Price " _
          & "FROM PartTable INNER JOIN PsitTable ON PartTable.PARTREF = PsitTable.PIPART INNER JOIN " _
          & "CihdTable ON PsitTable.PIPACKSLIP = CihdTable.INVPACKSLIP " _
          & "WHERE INVNO = ? "
   
   Set RdoQry2 = RdoCon.CreateQuery("", sSql)
   
   'Invoice Headers
   sSql = "SELECT INVNO,INVPRE,INVDATE,INVSHIPDATE,INVSTADR,INVARACCT,INVTAX,INVFREIGHT,CUQBNAME," _
          & "INVTAXACCT,INVTOTAL,INVPIF,INVTYPE,INVCOMMENTS,INVSO " _
          & "FROM CihdTable " _
          & "INNER JOIN CustTable ON CihdTable.INVCUST = CustTable.CUREF " _
          & "INNER JOIN PshdTable ON CihdTable.INVNO = PshdTable.PSINVOICE " _
          & "WHERE INVDATE >= '" & txtStart & "' AND INVDATE <='" & txtEnd & "'"
   bSqlRows = GetDataSet(RdoInv)
   
   If bSqlRows Then
      prg1.Visible = True
      prg1.Max = Val(lblFound)
      Open sFileName For Output As iFile
      
      With RdoInv
         sBuf = ""
         For i = 0 To UBound(sInvHdr)
            sBuf = sBuf & sInvHdr(i) & vbTab
         Next
         Print #iFile, sBuf
         
         sBuf = ""
         For i = 0 To UBound(sInvItHdr)
            sBuf = sBuf & sInvItHdr(i) & vbTab
         Next
         Print #iFile, sBuf
         Print #iFile, sQBEndTrans
         
         While Not .EOF
            If !INVPIF = 0 Then sPIF = "N" Else sPIF = "Y"
            If !INVTAX = 0 Then sTaxable = "N" Else sTaxable = "Y"
            
            sTerms = ""
            sDue = ""
            
            Select Case !INVTYPE
               Case "SO"
                  ' Invoiced From Sales Order
                  RdoQry1(0) = !INVNO
                  bSqlRows = GetQuerySet(RdoIt, RdoQry1)
                  sSql = "SELECT INVDAYS,INVARDISC,INVNETDAYS FROM CihdTable WHERE INVNO = " & !INVNO
                  bSqlRows = GetDataSet(RdoTerms)
                  If bSqlRows Then
                     sTerms = "" & RdoTerms!INVARDISC & "% " & RdoTerms!INVDAYS & " Net " & RdoTerms!INVNETDAYS
                     sDue = Format(DateAdd("d", Val("" & RdoTerms!INVNETDAYS), !INVDATE), "mm/dd/yyyy")
                  End If
                  bNoItems = False
               Case "PS"
                  ' Invoiced From Packslip
                  RdoQry2(0) = !INVNO
                  bSqlRows = GetQuerySet(RdoIt, RdoQry2)
                  
                  sSql = "SELECT SODAYS,SOARDISC,SONETDAYS FROM SohdTable WHERE SONUMBER = " & !INVSO
                  bSqlRows = GetDataSet(RdoTerms)
                  If bSqlRows Then
                     sTerms = "" & RdoTerms!INVARDISC & "% " & RdoTerms!INVDAYS & " Net " & RdoTerms!INVNETDAYS
                     sDue = Format(DateAdd("d", Val("" & RdoTerms!SONETDAYS), !INVDATE), "mm/dd/yyyy")
                  End If
                  bNoItems = False
               Case Else
                  ' Credit/Debit Memo
                  bNoItems = True
            End Select
            Set RdoTerms = Nothing
            
            'Output invoice header
            ' * Not outputed from ES/2002
            sBuf = Right(sInvHdr(0), Len(sInvHdr(0)) - 1) & vbTab
            sBuf = sBuf & "" & vbTab
            sBuf = sBuf & "INVOICE" & vbTab
            sBuf = sBuf & "" & Format(Trim(!INVDATE), "m/d/yyyy") & vbTab
            sBuf = sBuf & "" & Trim(txtQBAR) & vbTab
            sBuf = sBuf & "" & Trim(!CUQBNAME) & vbTab '
            sBuf = sBuf & "" & vbTab
            sBuf = sBuf & "" & Trim(!INVTOTAL) & vbTab
            sBuf = sBuf & "" & !INVNO & vbTab
            sBuf = sBuf & "" & Trim(!INVCOMMENTS) & vbTab
            sBuf = sBuf & "" & vbTab
            sBuf = sBuf & "" & vbTab
            sBuf = sBuf & "" & sTaxable & vbTab
            sBuf = sBuf & "" & Trim(sAddr(0)) & vbTab
            sBuf = sBuf & "" & Trim(sAddr(1)) & vbTab
            sBuf = sBuf & "" & Trim(sAddr(2)) & vbTab
            sBuf = sBuf & "" & Trim(sAddr(3)) & vbTab
            sBuf = sBuf & "" & Trim(sAddr(4)) & vbTab
            sBuf = sBuf & "" & sDue & vbTab
            sBuf = sBuf & "" & sTerms & vbTab
            sBuf = sBuf & "" & sPIF & vbTab
            sBuf = sBuf & "" & vbTab
            sBuf = sBuf & "" & Format(!INVSHIPDATE, "m/d/yyyy") & vbTab
            Print #iFile, sBuf
            
            ' Output invoice items
            ' * Not outputed from ES/2002
            If bNoItems = False Then
               While Not RdoIt.EOF
                  sBuf = Right(sInvItHdr(0), Len(sInvItHdr(0)) - 1) & vbTab 'ID
                  sBuf = sBuf & "" & vbTab 'SPLID
                  sBuf = sBuf & "INVOICE" & vbTab 'Type
                  sBuf = sBuf & "" & Trim(!INVDATE) & vbTab 'Date
                  sBuf = sBuf & "" & txtQBSales & vbTab 'Accnt
                  sBuf = sBuf & "" & !CUQBNAME & vbTab 'Customer
                  sBuf = sBuf & "" & vbTab 'Class *
                  sBuf = sBuf & "" & (RdoIt!Price * RdoIt!Qty) * -1 & vbTab 'Amount
                  sBuf = sBuf & "" & vbTab 'Docnum *
                  sBuf = sBuf & "" & Trim(RdoIt!Desciption) & vbTab 'Memo *
                  sBuf = sBuf & "" & vbTab 'Clear *
                  sBuf = sBuf & "" & (RdoIt!Qty) * -1 & vbTab 'Qty
                  sBuf = sBuf & "" & RdoIt!Price & vbTab 'Price
                  sBuf = sBuf & "" & RdoIt!Part & vbTab 'Inv Item
                  sBuf = sBuf & "" & vbTab
                  sBuf = sBuf & "N" & vbTab 'Taxable
                  sBuf = sBuf & "N" & vbTab 'VALADJ *
                  sBuf = sBuf & "NOTHING" & vbTab 'REIMBEXP
                  sBuf = sBuf & "" & vbTab 'SERVICEDATE *
                  sBuf = sBuf & "" & vbTab 'EXTRA *
                  Print #iFile, sBuf
                  RdoIt.MoveNext
               Wend
            End If
            
            If !INVFREIGHT <> 0 Then
               sBuf = Right(sInvItHdr(0), Len(sInvItHdr(0)) - 1) & vbTab
               sBuf = sBuf & vbTab
               sBuf = sBuf & "INVOICE" & vbTab
               sBuf = sBuf & "" & Trim(!INVDATE) & vbTab
               sBuf = sBuf & "" & txtFrtAcct & vbTab
               sBuf = sBuf & "" & Trim(!CUQBNAME) & vbTab
               sBuf = sBuf & "" & vbTab
               sBuf = sBuf & "" & (!INVFREIGHT * -1) & vbTab
               sBuf = sBuf & "" & vbTab
               sBuf = sBuf & "" & "Invoice Freight" & vbTab
               sBuf = sBuf & "N" & vbTab
               sBuf = sBuf & "" & vbTab
               sBuf = sBuf & "" & !INVFREIGHT & vbTab
               sBuf = sBuf & "" & Trim(txtFrtItm) & vbTab
               sBuf = sBuf & "" & vbTab
               sBuf = sBuf & "N" & vbTab
               sBuf = sBuf & "N" & vbTab
               sBuf = sBuf & "NOTHING" & vbTab
               sBuf = sBuf & "" & vbTab
               sBuf = sBuf & "" & vbTab
               Print #iFile, sBuf
            End If
            
            sBuf = Right(sInvItHdr(0), Len(sInvItHdr(0)) - 1) & vbTab
            sBuf = sBuf & vbTab
            sBuf = sBuf & "INVOICE" & vbTab
            sBuf = sBuf & "" & Trim(!INVDATE) & vbTab
            sBuf = sBuf & "" & txtQBTax & vbTab
            sBuf = sBuf & "" & txtTaxVend & vbTab
            sBuf = sBuf & "" & vbTab
            sBuf = sBuf & "" & (!INVTAX * -1) & vbTab
            sBuf = sBuf & "" & vbTab
            sBuf = sBuf & "" & vbTab
            sBuf = sBuf & "N" & vbTab
            sBuf = sBuf & "" & vbTab
            sBuf = sBuf & "" & vbTab
            sBuf = sBuf & "" & txtTaxItem & vbTab
            sBuf = sBuf & "" & vbTab
            sBuf = sBuf & "N" & vbTab
            sBuf = sBuf & "N" & vbTab
            sBuf = sBuf & "NOTHING" & vbTab
            sBuf = sBuf & "" & vbTab
            sBuf = sBuf & "AUTOSTAX" & vbTab
            Print #iFile, sBuf
            
            Print #iFile, Right(sQBEndTrans, Len(sQBEndTrans) - 1)
            Set RdoIt = Nothing
            prg1.Value = prg1.Value + 1
            iLastInv = !INVNO
            sLastInvPre = !INVPRE
            .MoveNext
         Wend
      End With
      
      ' Remeber last invoice and last date
      rdoQB.Edit
      rdoQB!QBLASTARINV = iLastInv
      rdoQB!QBLASTAREXP = Format(Now, "mm/dd/yyyy")
      rdoQB.Update
      lblDte = Format("" & Trim(rdoQB!QBLASTAREXP), "m/d/yyyy")
      lblLast = sLastInvPre & Format(iLastInv, "000000")
      
      
      smsg = "Successfully Exported AR Activity."
      Sysmsg smsg, True
   Else
      smsg = "Could Not Build QuickBooks" & vbCrLf & "AR Export."
      MsgBox smsg, vbExclamation, Caption
   End If
   
   CleanUp:
   prg1.Value = 0
   prg1.Visible = False
   Close iFile
   Set RdoInv = Nothing
   Set RdoIt = Nothing
   Set RdoQry1 = Nothing
   Set RdoQry2 = Nothing
   Set RdoQry3 = Nothing
   Set RdoQry4 = Nothing
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
   txtStart = CheckDate(txtStart)
   GetInvoices
End Sub

Public Function GetQBSettings() As Byte
   Dim rdoPre As rdoResultset
   
   sSql = "SELECT * FROM QkbkTable"
   
   bSqlRows = GetDataSet(rdoQB, ES_KEYSET)
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
         
         sSql = "SELECT INVPRE FROM CihdTable WHERE INVNO = " & !QBLASTARINV
         bSqlRows = GetDataSet(rdoPre)
         If bSqlRows Then
            lblLast = rdoPre!INVPRE & Format("" & Trim(!QBLASTARINV), "000000")
         End If
         Set rdoPre = Nothing
      End With
      GetQBSettings = 1
   Else
      GetQBSettings = 0
   End If
End Function

Public Sub IniQB()
   On Error GoTo DiaErr1
   
   bGoodQB = GetQBSettings
   
   ' first time
   If bGoodQB < 1 Then
      sSql = "INSERT INTO QkBkTable(QBREF) VALUES(1)"
      RdoCon.Execute sSql, rdExecDirect
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
      rdoQB.Edit
      rdoQB!QBTAXITEM = Trim(txtTaxItem)
      rdoQB.Update
      If Err > 0 Then ValidateEdit Me
   End If
End Sub

Private Sub txtTaxVend_LostFocus()
   txtTaxVend = CheckLen(txtTaxVend, 30)
   If bGoodQB Then
      On Error Resume Next
      rdoQB.Edit
      rdoQB!QBTAXVEND = Trim(txtTaxVend)
      rdoQB.Update
      If Err > 0 Then ValidateEdit Me
   End If
End Sub
