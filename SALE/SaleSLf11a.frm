VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Begin VB.Form SaleSLf11a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Create Sales Orders From Estimates"
   ClientHeight    =   4380
   ClientLeft      =   1845
   ClientTop       =   1065
   ClientWidth     =   6165
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   HelpContextID   =   2101
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtCmt 
      Height          =   975
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   31
      Tag             =   "9"
      ToolTipText     =   "Comments (3072 Chars Max)"
      Top             =   4200
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.TextBox txtUnitPrc 
      Height          =   285
      Left            =   4680
      TabIndex        =   30
      Tag             =   "1"
      ToolTipText     =   "Enter Quantity"
      Top             =   4200
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "SaleSLf11a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   29
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.TextBox txtQty 
      Height          =   285
      Left            =   3960
      TabIndex        =   5
      Tag             =   "1"
      ToolTipText     =   "Enter Quantity"
      Top             =   3480
      Width           =   1095
   End
   Begin VB.ComboBox txtDue 
      Height          =   315
      Left            =   1680
      TabIndex        =   4
      Tag             =   "4"
      Top             =   3480
      Width           =   1215
   End
   Begin VB.ComboBox cmbBid 
      Height          =   288
      Left            =   1920
      TabIndex        =   3
      ToolTipText     =   "Contains Accepted Estimates"
      Top             =   2280
      Width           =   975
   End
   Begin VB.TextBox txtCpo 
      Height          =   285
      Left            =   1680
      TabIndex        =   6
      Tag             =   "3"
      ToolTipText     =   "Customer PO (Optional) 20 Char Max"
      Top             =   3840
      Width           =   2085
   End
   Begin VB.CommandButton cmdNte 
      DisabledPicture =   "SaleSLf11a.frx":07AE
      DownPicture     =   "SaleSLf11a.frx":1120
      Height          =   315
      Left            =   3440
      Picture         =   "SaleSLf11a.frx":1A92
      Style           =   1  'Graphical
      TabIndex        =   15
      TabStop         =   0   'False
      ToolTipText     =   "Selling And Collection Notes"
      Top             =   960
      Width           =   375
   End
   Begin VB.CommandButton cmdVew 
      DownPicture     =   "SaleSLf11a.frx":2094
      Height          =   315
      Left            =   3000
      Picture         =   "SaleSLf11a.frx":256E
      Style           =   1  'Graphical
      TabIndex        =   14
      TabStop         =   0   'False
      ToolTipText     =   "View Sales Order List"
      Top             =   960
      Width           =   375
   End
   Begin VB.Timer tmr1 
      Enabled         =   0   'False
      Interval        =   30000
      Left            =   5760
      Top             =   5640
   End
   Begin VB.TextBox txtSon 
      Height          =   285
      Left            =   2160
      TabIndex        =   1
      Tag             =   "1"
      ToolTipText     =   "Enter New Sales Order Number"
      Top             =   960
      Width           =   735
   End
   Begin VB.ComboBox cmbCst 
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1680
      TabIndex        =   2
      Tag             =   "3"
      ToolTipText     =   "Customers With Qualifying Estimates. Select Customer From List"
      Top             =   1440
      Width           =   1555
   End
   Begin VB.ComboBox cmbPre 
      Height          =   315
      Left            =   1680
      TabIndex        =   0
      Tag             =   "3"
      Text            =   "S"
      ToolTipText     =   "Select or Enter Type A thru Z"
      Top             =   960
      Width           =   520
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   315
      Left            =   5160
      TabIndex        =   7
      ToolTipText     =   "Add This New Sales Order"
      Top             =   3840
      Width           =   915
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5160
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   0
      Width           =   915
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5280
      Top             =   5640
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4380
      FormDesignWidth =   6165
   End
   Begin VB.Label lblClass 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   3120
      TabIndex        =   28
      ToolTipText     =   "Estimate Type"
      Top             =   2280
      Width           =   612
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity"
      Height          =   255
      Index           =   9
      Left            =   3000
      TabIndex        =   27
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date Due"
      Height          =   252
      Index           =   8
      Left            =   240
      TabIndex        =   26
      Top             =   3480
      Width           =   1332
   End
   Begin VB.Label lblQuantity 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   3960
      TabIndex        =   25
      ToolTipText     =   "Bid Quantity"
      Top             =   3120
      Width           =   1092
   End
   Begin VB.Label lblDate 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1680
      TabIndex        =   24
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity"
      Height          =   255
      Index           =   7
      Left            =   3000
      TabIndex        =   23
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Bid Date"
      Height          =   252
      Index           =   6
      Left            =   240
      TabIndex        =   22
      Top             =   3120
      Width           =   1332
   End
   Begin VB.Label lblPre 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1680
      TabIndex        =   21
      ToolTipText     =   "Estimate Prefix"
      Top             =   2280
      Width           =   252
   End
   Begin VB.Label lblPart 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   1680
      TabIndex        =   20
      ToolTipText     =   "Bid Part Number"
      Top             =   2640
      Width           =   3372
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      Height          =   252
      Index           =   5
      Left            =   240
      TabIndex        =   19
      Top             =   2640
      Width           =   1332
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Estimates"
      Height          =   252
      Index           =   4
      Left            =   240
      TabIndex        =   18
      Top             =   2280
      Width           =   1332
   End
   Begin VB.Label lblNotice 
      BackStyle       =   0  'Transparent
      Caption         =   "Note: The Last Sales Order Number Has Changed"
      Height          =   252
      Left            =   240
      TabIndex        =   17
      Top             =   300
      Visible         =   0   'False
      Width           =   4332
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Order"
      Height          =   252
      Index           =   1
      Left            =   240
      TabIndex        =   16
      Top             =   3840
      Width           =   1332
   End
   Begin VB.Label txtNme 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1680
      TabIndex        =   13
      ToolTipText     =   "Customer Name"
      Top             =   1800
      Width           =   3375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer"
      Height          =   252
      Index           =   2
      Left            =   240
      TabIndex        =   12
      Top             =   1440
      Width           =   1332
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "SO Number"
      Height          =   252
      Index           =   0
      Left            =   240
      TabIndex        =   11
      Top             =   960
      Width           =   1452
   End
   Begin VB.Label lblLst 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1680
      TabIndex        =   10
      ToolTipText     =   "Last Sales Order Entered"
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Last Sales Order"
      Height          =   252
      Index           =   3
      Left            =   240
      TabIndex        =   9
      Top             =   600
      Width           =   1452
   End
End
Attribute VB_Name = "SaleSLf11a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'New 4/14/06 (created from SaleSLE01a)
Option Explicit
Dim bCutOff As Byte
Dim bOnLoad As Byte
Dim bSoAdded As Byte
Dim bSoExists As Byte
Dim bGoodBid As Byte
Dim bGoodCust As Byte

Dim iDays As Integer
Dim iFrtDays As Integer
Dim iNetDays As Integer

Dim sLastPrefix As String
Dim sNewsonumber As String
Dim sCust As String
Dim sStName As String
Dim sStAdr As String
Dim sContact As String
Dim sConIntPhone As String
Dim sConPhone As String
Dim sConIntFax As String
Dim sConFax As String
Dim sConExt As String
Dim sDivision As String
Dim sOldSoNumber As String
Dim sRegion As String
Dim sSterms As String
Dim sVia As String
Dim sFob As String
Dim sSlsMan As String
Dim sTaxExempt As String

Dim cDiscount As Currency

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd


Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   lblNotice.ForeColor = ES_RED
   txtDue = Format(GetServerDateTime, "mm/dd/yyyy")
   
End Sub


Private Sub cmbBid_Click()
   bGoodBid = GetThisBid()
   
End Sub


Private Sub cmbBid_LostFocus()
   Dim bByte As Byte
   Dim iList As Integer
   On Error Resume Next
   cmbBid = Format(Abs(Val(cmbBid)), "000000")
   If cmbBid.ListCount = 0 Then Exit Sub
   For iList = 0 To cmbBid.ListCount - 1
      If cmbBid = cmbBid.List(iList) Then bByte = 1
   Next
   If bByte = 0 Then
      MsgBox "That Estimate Does Not Qualify.", _
         vbInformation, Caption
      cmbBid = cmbBid.List(0)
   Else
      bGoodBid = GetThisBid(1)
   End If
   
End Sub


Private Sub cmbCst_Change()
   CloseBoxes 1
   
End Sub

Private Sub cmbCst_Click()
   FindCustomer Me, cmbCst, False
   FillEstimates
   
End Sub

Private Sub cmbCst_LostFocus()
   Dim bByte As Byte
   Dim iList As Integer
   cmbCst = CheckLen(cmbCst, 10)
   If cmbCst.ListCount > 0 Then
      For iList = 0 To cmbCst.ListCount - 1
         If cmbCst = cmbCst.List(iList) Then bByte = 1
      Next
      If bByte = 0 Then
         MsgBox "Please Select A Customer From The List.", _
            vbInformation, Caption
         cmbCst = cmbCst.List(0)
      End If
      FindCustomer Me, cmbCst, False
      FillEstimates
      lblNotice.Visible = False
   End If
   
End Sub


Private Sub cmbPre_LostFocus()
   Dim a As Integer
   cmbPre = CheckLen(cmbPre, 1)
   On Error Resume Next
   a = Asc(Left(cmbPre, 1))
   If a < 65 Or a > 90 Then
      MsgBox "Must Be Between A and Z..", vbInformation, Caption
      cmbPre = sLastPrefix
   End If
   If Len(Trim(cmbPre)) = 0 Then cmbPre = sLastPrefix
   
End Sub


Private Sub cmdAdd_Click()
   Dim bByte As Byte
   Dim lNewSoNum As Long
   lNewSoNum = Val(txtSon)
   lblNotice.Visible = False
   If lNewSoNum > 99999 Then
      MsgBox "Sales Orders May Be In The Range 1 to 99999.", _
         vbInformation, Caption
      Exit Sub
   End If
   If txtNme.ForeColor <> ES_RED Then
      bSoExists = GetOldSalesOrder()
      If bSoExists Then
         MsgBox "That Sales Order Was Previously Recorded.", _
            vbInformation, Caption
         txtSon = sNewsonumber
         On Error Resume Next
         txtSon.SetFocus
      Else
         bByte = CheckCustomerPO
         If bByte = 1 Then
            bByte = MsgBox("The Customer PO Is In Use. Continue?", _
                    ES_YESQUESTION, Caption)
            If bByte = vbNo Then
               CancelTrans
               Exit Sub
            End If
         End If
         tmr1.Enabled = False
         AddSalesOrder
      End If
   Else
      MsgBox "Requires A Valid Customer.", _
         vbExclamation, Caption
   End If
   
End Sub

Private Sub cmdCan_Click()
   tmr1.Enabled = False
   sLastPrefix = cmbPre
   Unload Me
   
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   tmr1.Enabled = False
   
End Sub


Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 2161
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub

Private Sub cmdNte_Click()
   SaleSLe02d.lblCustomer = cmbCst
   SaleSLe02d.Show
   
End Sub

Private Sub cmdVew_Click()
   Dim iList As Integer
   Dim iCol As Integer
   Dim iRows As Integer
   Dim RdoVew As ADODB.Recordset
   On Error Resume Next
   iRows = 10
   With SalesSoView.Grd
      .Rows = iRows
      .ColAlignment(0) = 0
      .ColAlignment(1) = 0
      If Screen.Width > 9999 Then
         .ColWidth(0) = 1100 * 1.25
         .ColWidth(1) = 1550 * 1.25
         .ColWidth(2) = 1900 * 1.25
      Else
         .ColWidth(0) = 1100
         .ColWidth(1) = 1550
         .ColWidth(2) = 1900
      End If
   End With
   sSql = "SELECT SONUMBER,SOTYPE,SOCUST,SODATE,CUREF,CUNICKNAME FROM " _
          & "SohdTable,CustTable WHERE (SOCUST=CUREF) " _
          & "ORDER BY SONUMBER DESC"
   MouseCursor 13
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoVew)
   If bSqlRows Then
      With RdoVew
         Do Until .EOF
            iList = iList + 1
            If iList > 300 Then Exit Do
            iRows = iRows + 1
            SalesSoView.Grd.Rows = iRows
            SalesSoView.Grd.Row = iRows - 11
            
            SalesSoView.Grd.Col = 0
            SalesSoView.Grd = "" & Trim(!SOTYPE) & Format(!SoNumber, SO_NUM_FORMAT)
            
            SalesSoView.Grd.Col = 1
            SalesSoView.Grd = "" & Format(!SODATE, "mm/dd/yy")
            
            SalesSoView.Grd.Col = 2
            SalesSoView.Grd = Trim(!CUNICKNAME)
            
            .MoveNext
         Loop
         ClearResultSet RdoVew
      End With
      If iList > 9 Then SalesSoView.Grd.Rows = iList + 1
   End If
   MouseCursor 0
   Set RdoVew = Nothing
   On Error GoTo 0
   SalesSoView.Show
   
End Sub

Private Sub Form_Activate()
   Dim b As Byte
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillCombo
      GetLastSalesOrder sOldSoNumber, sNewsonumber, True
      If cmbCst.ListCount > 0 Then FillEstimates
      bSoAdded = 0
      tmr1.Enabled = True
      lblPre.Height = cmbBid.Height
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   Dim iChar As Integer
   FormLoad Me
   FormatControls
   sLastPrefix = GetSetting("Esi2000", "EsiSale", "LastPrefix", sLastPrefix)
   If Len(sLastPrefix) = 0 Then sLastPrefix = "S"
   cmbPre = sLastPrefix
   For iChar = 65 To 90
      AddComboStr cmbPre.hWnd, Chr$(iChar)
   Next
   bOnLoad = 1
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveSetting "Esi2000", "EsiSale", "LastPrefix", sLastPrefix
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   cUR.CurrentCustomer = cmbCst
   FormUnload
   If txtNme.ForeColor <> ES_RED Then SaveCurrentSelections
   Set SaleSLf11a = Nothing
   
End Sub


Private Sub lblLst_Change()
   If sOldSoNumber <> "" And sOldSoNumber <> lblLst Then _
      lblNotice.Visible = True
   
End Sub


Private Sub tmr1_Timer()
   If Val(txtSon) > Val(Right(lblLst, 5)) + 1 Then
      GetLastSalesOrder sOldSoNumber, sNewsonumber, False
   Else
      GetLastSalesOrder sOldSoNumber, sNewsonumber, True
   End If
   
End Sub


Private Sub txtCpo_LostFocus()
   Dim bByte As Byte
   txtCpo = CheckLen(txtCpo, 20)
   bByte = CheckCustomerPO()
   If bByte = 1 Then
      bByte = MsgBox("That Customer PO Is In Use. Proceed Anyway?", _
              ES_NOQUESTION, Caption)
      If bByte = vbNo Then txtCpo = ""
   End If
   
End Sub


Private Sub txtDue_DropDown()
   ShowCalendarEx Me
   
End Sub

Private Sub txtDue_LostFocus()
   txtDue = CheckDateEx(txtDue)
   
End Sub


Private Sub txtNme_Change()
   If txtNme = "*** Customer Wasn't Found ***" Then
      txtNme.ForeColor = ES_RED
   Else
      txtNme.ForeColor = Es_TextForeColor
   End If
   
End Sub

Private Sub txtQty_LostFocus()
   txtQty = Format(Abs(Val(txtQty)), "######0.000")
   If Val(txtQty) = 0 Then
      Beep
      txtQty = lblQuantity
   End If
   If Val(txtQty) <> Val(lblQuantity) Then _
          MsgBox "The SO Quantity Does Not Match The Estimate Quantity.", _
          vbInformation, Caption
   
End Sub


Private Sub txtSon_Click()
   txtSon_GotFocus
   
End Sub

Private Sub txtSon_GotFocus()
   tmr1.Enabled = True
   
End Sub


Private Sub txtSon_LostFocus()
   txtSon = CheckLen(txtSon, SO_NUM_SIZE)
   txtSon = Format(Abs(Val(txtSon)), SO_NUM_FORMAT)
   If Val(txtSon) > Val(Right(lblLst, SO_NUM_SIZE)) Then lblNotice.Visible = False
   
End Sub

Private Sub AddSalesOrder()
   Dim bResponse As Byte
   Dim sNewDate As Variant
   Dim strComt As String
   
   bGoodCust = GetCustomerData()
   If bCutOff = 1 Then
      MsgBox "This Customer's Credit Is On Hold.", _
         vbInformation, Caption
      bGoodCust = 0
   End If
   If Not bGoodCust Then Exit Sub
   On Error GoTo DiaErr1
   If Len(txtSon) < SO_NUM_SIZE Then
      txtSon = Format(Abs(Val(txtSon)), SO_NUM_FORMAT)
   End If
   sNewDate = Format(ES_SYSDATE, "mm/dd/yyyy")
   clsADOCon.BeginTrans
   clsADOCon.ADOErrNum = 0
'   sSql = "INSERT SohdTable (SONUMBER,SOTYPE,SOCUST,SODATE," _
'          & "SOSALESMAN,SOSTNAME,SOSTADR,SODIVISION,SOREGION,SOSTERMS," _
'          & "SOVIA,SOFOB,SOARDISC,SODAYS,SONETDAYS,SOFREIGHTDAYS," _
'          & "SOTEXT,SOTAXEXEMPT,SOPO,SOFROMBID) " _
'          & "VALUES(" & Val(txtSon) & ",'" & cmbPre & "','" _
'          & sCust & "','" & sNewDate & "','" & sSlsMan & "','" _
'          & sStName & "','" & sStAdr & "','" & sDivision & "','" _
'          & sRegion & "','" & sSterms & "','" & sVia & "','" _
'          & sFob & "'," & cDiscount & "," & iDays & "," & iNetDays _
'          & "," & iFrtDays & ",'" & txtSon & "','" & sTaxExempt & "','" _
'          & Trim(txtCpo) & "'," & Val(cmbBid) & ")"
   sSql = "INSERT SohdTable (SONUMBER,SOTYPE,SOCUST,SODATE," _
          & "SOSALESMAN,SOSTNAME,SOSTADR,SODIVISION,SOREGION,SOSTERMS," _
          & "SOVIA,SOFOB,SOARDISC,SODAYS,SONETDAYS,SOFREIGHTDAYS," _
          & "SOTAXEXEMPT,SOPO,SOFROMBID) " _
          & "VALUES(" & Val(txtSon) & ",'" & cmbPre & "','" _
          & sCust & "','" & sNewDate & "','" & sSlsMan & "','" _
          & sStName & "','" & sStAdr & "','" & sDivision & "','" _
          & sRegion & "','" & sSterms & "','" & sVia & "','" _
          & sFob & "'," & cDiscount & "," & iDays & "," & iNetDays _
          & "," & iFrtDays & ",'" & sTaxExempt & "','" _
          & Trim(txtCpo) & "'," & Val(cmbBid) & ")"
   clsADOCon.ExecuteSql sSql 'rdExecDirect
   
   strComt = ReplaceSingleQuote(Trim(txtCmt))
   
   sSql = "INSERT SoitTable (ITSO,ITNUMBER,ITPART,ITQTY,ITSCHED," _
          & "ITBOOKDATE,ITUSER, ITDOLLORIG, ITCOMMENTS)" _
          & "VALUES(" & Val(txtSon) & ",1,'" _
          & Compress(lblPart) & "'," & Val(txtQty) & ",'" & txtDue & "','" _
          & txtDue & "','" & sInitials & "','" & Val(txtUnitPrc) & "','" & strComt & "')"
   clsADOCon.ExecuteSql sSql 'rdExecDirect
   
   sSql = "UPDATE EstiTable SET BIDSONUMBER=" & Val(txtSon) & " " _
          & "WHERE BIDREF=" & Val(cmbBid) & " "
   clsADOCon.ExecuteSql sSql 'rdExecDirect
   
   If clsADOCon.RowsAffected Then
      bSoAdded = 1
      On Error Resume Next
      clsADOCon.CommitTrans
      MsgBox "Sales Order From Estimate Was Added.", _
         vbInformation, Caption
      sSql = "UPDATE SohdTable SET SOCCONTACT='" & sContact & "'," _
             & "SOCPHONE='" & sConPhone & "',SOCINTPHONE='" _
             & sConIntPhone & "',SOCINTFAX='" & sConIntFax _
             & "',SOCFAX='" & sConFax & "',SOCEXT=" & sConExt _
             & " WHERE SONUMBER=" & Val(txtSon) & " "
      clsADOCon.ExecuteSql sSql 'rdExecDirect
      
      sSql = "UPDATE ComnTable SET COLASTSALESORDER='" & Trim(cmbPre) _
             & Trim(txtSon) & "' WHERE COREF=1"
      clsADOCon.ExecuteSql sSql 'rdExecDirect
      
      bResponse = MsgBox("The Sales Order Is Complete. Edit It?", _
                  ES_YESQUESTION, Caption)
      If bResponse = vbYes Then
         Unload Me
         'MM 2/15/2009 moved the Unload Me ahead of show form.
         SetActiveWindow (SaleSLe02a.hWnd)
         SaleSLe02a.Show
         SaleSLe02a.cmbSon.SetFocus
      Else
         On Error Resume Next
         lblLst = txtSon
         CloseBoxes
         FillCombo
         GetLastSalesOrder sOldSoNumber, sNewsonumber, True
         txtSon.SetFocus
      End If
   Else
      clsADOCon.RollbackTrans
      MsgBox "Couldn't Add Sales Order.", vbExclamation, Caption
   End If
   Exit Sub
   
DiaErr1:
   If Left(Err.Description, 5) = "01000" Then
      MsgBox "Sales Order Number " & txtSon & " Was Recently Used By  " & vbCrLf _
         & "Another Process. Please Select The Next Number.", _
         vbInformation, Caption
      GetLastSalesOrder sOldSoNumber, sNewsonumber, True
      sOldSoNumber = lblLst
      lblNotice.Visible = False
      tmr1.Enabled = True
   Else
      sProcName = "addsaleso"
      CurrError.Number = Err.Number
      CurrError.Description = Err.Description
      DoModuleErrors Me
   End If
   
End Sub

Private Function GetCustomerData() As Byte
   Dim RdoCst As ADODB.Recordset
   sCust = Compress(cmbCst)
   On Error GoTo DiaErr1
   sSql = "SELECT CUREF,CUSTNAME,CUSTNAME,CUSTADR,CUARDISC," _
          & "CUDAYS,CUNETDAYS,CUDIVISION,CUREGION,CUSTERMS," _
          & "CUVIA,CUFOB,CUSALESMAN,CUDISCOUNT,CUSTSTATE," _
          & "CUSTCITY,CUSTZIP,CUCCONTACT,CUCPHONE,CUCEXT,CUCINTPHONE," _
          & "CUFRTDAYS,CUINTFAX,CUFAX,CUTAXEXEMPT,CUCUTOFF " _
          & "FROM CustTable WHERE CUREF='" & sCust & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCst)
   If bSqlRows Then
      With RdoCst
         bCutOff = !CUCUTOFF
         sStName = "" & Trim(!CUSTNAME)
         sStAdr = "" & Trim(!CUSTADR) & vbCrLf _
                  & "" & Trim(!CUSTCITY) & ", " & Trim(!CUSTSTATE) _
                  & "  " & Trim(!CUSTZIP)
         sDivision = "" & Trim(!CUDIVISION)
         sRegion = "" & Trim(!CUREGION)
         sSterms = "" & Trim(!CUSTERMS)
         sVia = "" & Trim(!CUVIA)
         sFob = "" & Trim(!CUFOB)
         sSlsMan = "" & Trim(!CUSALESMAN)
         sContact = "" & Trim(!CUCCONTACT)
         sConIntPhone = "" & Trim(!CUCINTPHONE)
         sConPhone = "" & Trim(!CUCPHONE)
         sConIntFax = "" & Trim(!CUINTFAX)
         sConFax = "" & Trim(!CUFAX)
         sConExt = "" & Trim(str$(!CUCEXT))
         cDiscount = Format(0 + !CUARDISC, "##0.000")
         iDays = Format(!CUDAYS, "###0")
         iNetDays = Format(!CUNETDAYS, "###0")
         iFrtDays = Format(!CUFRTDAYS, "##0")
         sTaxExempt = "" & Trim(!CUTAXEXEMPT)
         ClearResultSet RdoCst
      End With
      GetCustomerData = True
   Else
      sStName = ""
      sStAdr = ""
      sDivision = ""
      sRegion = ""
      sSterms = ""
      sVia = ""
      sFob = ""
      sSlsMan = ""
      iFrtDays = 0
      MsgBox "Couldn't Retrieve Customer.", vbExclamation, Caption
      GetCustomerData = False
   End If
   Set RdoCst = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getcustda"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function


Private Sub FillEstimates()
   CloseBoxes 1
   On Error GoTo DiaErr1
   sSql = "SELECT BIDREF,BIDNUM,BIDPRE,BIDPART,BIDCUST,BIDQUANTITY," _
          & "BIDDATE, BIDCOMPLETE,CUREF,PARTREF,PARTNUM " _
          & "FROM EstiTable,CustTable,PartTable WHERE (BIDCUST=CUREF " _
          & "AND BIDPART=PARTREF AND BIDACCEPTED=1 " _
          & "AND BIDSONUMBER=0 AND BIDCUST='" & Compress(cmbCst) & "') "
   LoadComboBox cmbBid
   If cmbBid.ListCount > 0 Then
      bGoodBid = GetThisBid()
   Else
      bGoodBid = 0
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "fillestima"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Function GetThisBid(Optional SetTxtDue As Byte) As Byte
   CloseBoxes
   Dim RdoBid As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT BIDREF,BIDPRE,BIDCLASS,BIDDATE,BIDPART,BIDQUANTITY," _
          & "BIDUNITPRICE, BIDCOMMENT, PARTREF,PARTNUM FROM EstiTable,PartTable WHERE " _
          & "(BIDPART=PARTREF AND BIDREF=" & Val(cmbBid) & ")"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoBid, ES_FORWARD)
   If bSqlRows Then
      With RdoBid
         lblPre = "" & Trim(!BIDPRE)
         lblClass = "" & Trim(!BIDCLASS)
         lblPart = "" & Trim(!PartNum)
         lblDate = Format(!BIDDATE, "mm/dd/yyyy")
         lblQuantity = Format(!BIDQUANTITY, "######0.000")
         txtQty = Format(!BIDQUANTITY, "######0.000")
         txtUnitPrc = Format(Abs(Val(!BIDUNITPRICE)), "####0.00")
         txtCmt = "" & Trim(!BIDCOMMENT)
         
         GetThisBid = 1
         If bOnLoad = 0 Then
            txtDue.Enabled = True
            txtCpo.Enabled = True
            txtQty.Enabled = True
            cmdAdd.Enabled = True
         End If
      End With
      ClearResultSet RdoBid
   Else
      GetThisBid = 0
   End If
   On Error Resume Next
   If GetThisBid = 1 And SetTxtDue = 1 Then txtDue.SetFocus
   Set RdoBid = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getthisbid"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub CloseBoxes(Optional ClearCombo As Byte)
   If ClearCombo = 1 Then cmbBid.Clear
   bGoodBid = 0
   lblPart = ""
   lblDate = ""
   lblQuantity = ""
   txtCpo = ""
   lblClass = ""
   txtQty = "0.000"
   txtDue.Enabled = False
   txtCpo.Enabled = False
   txtQty.Enabled = False
   cmdAdd.Enabled = False
   
End Sub

Private Sub FillCombo()
   On Error GoTo DiaErr1
   cmbCst.Clear
   sSql = "SELECT DISTINCT BIDCUST,CUREF,CUNICKNAME FROM EstiTable," _
          & "CustTable WHERE (BIDCUST=CUREF AND BIDACCEPTED=1 AND " _
          & "BIDSONUMBER=0) ORDER BY BIDCUST"
   LoadComboBox cmbCst, 1
   If cmbCst.ListCount > 0 Then
      If cUR.CurrentCustomer <> "" Then cmbCst = cUR.CurrentCustomer
      FindCustomer Me, cmbCst, False
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub
