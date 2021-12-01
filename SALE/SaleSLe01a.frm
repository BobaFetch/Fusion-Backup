VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Begin VB.Form SaleSLe01a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Sales Order"
   ClientHeight    =   2655
   ClientLeft      =   1845
   ClientTop       =   1065
   ClientWidth     =   6495
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   HelpContextID   =   2101
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "SaleSLe01a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   17
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.TextBox txtCpo 
      Height          =   285
      Left            =   1680
      TabIndex        =   3
      Tag             =   "3"
      ToolTipText     =   "Customer PO (Optional) 20 Char Max"
      Top             =   2160
      Width           =   2085
   End
   Begin VB.CommandButton cmdNte 
      DisabledPicture =   "SaleSLe01a.frx":07AE
      DownPicture     =   "SaleSLe01a.frx":0C2B
      Height          =   315
      Left            =   3440
      Picture         =   "SaleSLe01a.frx":10A8
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Selling And Collection Notes"
      Top             =   960
      Width           =   375
   End
   Begin VB.CommandButton cmdVew 
      DownPicture     =   "SaleSLe01a.frx":1525
      Height          =   315
      Left            =   3000
      Picture         =   "SaleSLe01a.frx":19FF
      Style           =   1  'Graphical
      TabIndex        =   13
      TabStop         =   0   'False
      ToolTipText     =   "View Sales Order List"
      Top             =   960
      Width           =   375
   End
   Begin VB.CheckBox optNew 
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   2520
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CheckBox optRev 
      Caption         =   "Revise"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2880
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Timer tmr1 
      Enabled         =   0   'False
      Interval        =   30000
      Left            =   5880
      Top             =   1560
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
      ToolTipText     =   "Select Customer From List"
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
      Left            =   5520
      TabIndex        =   4
      ToolTipText     =   "Add This New Sales Order"
      Top             =   480
      Width           =   915
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5520
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   0
      Width           =   915
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5880
      Top             =   2040
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2655
      FormDesignWidth =   6495
   End
   Begin VB.Label lblNotice 
      Caption         =   "Note: The Last Sales Order Number Has Changed"
      Height          =   252
      Left            =   1680
      TabIndex        =   16
      Top             =   240
      Visible         =   0   'False
      Width           =   4092
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Order"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   15
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label txtNme 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1680
      TabIndex        =   10
      Top             =   1800
      Width           =   3375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   9
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "SO Number"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label lblLst 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1680
      TabIndex        =   7
      ToolTipText     =   "Last Sales Order Entered"
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Last Sales Order"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   1335
   End
End
Attribute VB_Name = "SaleSLe01a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'6/25/04 Added cmdNte
'7/15/05 Added COLASTSALESORDER
'10/14/05 Added a Sales Order Changed Notice and
'         softened the Duplicate Row Message
'11/23/05 Added Formuload if SO wasn't added
'1/11/06 Corrected Formunload if called from Revise
'4/14/06 Move some procedures to Mudule.EsiSale for new SaleSLf11a
Option Explicit
Dim bCutOff As Byte
Dim bOnLoad As Byte
Dim bSoAdded As Byte
Dim bSoExists As Byte
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
   
End Sub

Private Sub cmbCst_Click()
   FindCustomer Me, cmbCst, False
   
End Sub

Private Sub cmbCst_LostFocus()
   cmbCst = CheckLen(cmbCst, 10)
   FindCustomer Me, cmbCst, False
   lblNotice.Visible = False
   
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
   If lNewSoNum > 999999 Then
      MsgBox "Sales Orders May Be In The Range 1 to 999999.", _
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
      OpenHelpContext 2101
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
      GetLastSalesOrder sOldSoNumber, sNewsonumber, True
      FillCustomers
      If cUR.CurrentCustomer <> "" Then cmbCst = cUR.CurrentCustomer
      FindCustomer Me, cmbCst, False
      bSoAdded = 0
      tmr1.Enabled = True
      ' Removed at request of Dan Paquette 10/15/02
      '        sJournalId = GetOpenJournal("SJ", Format$(Now, "mm/dd/yy"))
      '            If Left(sJournalId, 4) = "None" Then
      '                sJournalId = ""
      '                b = 1
      '            Else
      '                If sJournalId = "" Then b = 0 Else b = 1
      '            End If
      '        If b = 0 Then
      '            MsgBox "There Is No Open Sales Journal For This Period.", _
      '                vbExclamation, Caption
      '            Sleep 500
      '            Unload Me
      '            Exit Sub
      '        End If
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   Dim iChar As Integer
   If optRev Then
      Move MdiSect.Left + 400, MdiSect.Top + 400
   Else
      FormLoad Me
   End If
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
   If optNew.Value = vbUnchecked And optRev.Value = vbUnchecked Then FormUnload
   If txtNme.ForeColor <> ES_RED Then SaveCurrentSelections
   If optRev = vbChecked Then SaleSLe02a.optNew = vbUnchecked
   Set SaleSLe01a = Nothing
   
End Sub

Private Sub lblLst_Change()
   If sOldSoNumber <> "" And sOldSoNumber <> lblLst Then _
      lblNotice.Visible = True
   
End Sub

Private Sub optRev_Click()
   'never visible-loaded from SaleSLe02a
   
End Sub

Private Sub tmr1_Timer()
'REMOVED 8/21/08 - CAUSING PROBLEMS AT LUMICOR
'   If Val(txtSon) > Val(Right(lblLst, 5)) + 1 Then
'      GetLastSalesOrder sOldSoNumber, sNewsonumber, False
'   Else
'      GetLastSalesOrder sOldSoNumber, sNewsonumber, True
'   End If
   
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


Private Sub txtNme_Change()
   If txtNme = "*** Customer Wasn't Found ***" Then
      txtNme.ForeColor = ES_RED
   Else
      txtNme.ForeColor = Es_TextForeColor
   End If
   
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



Private Function GetOldSalesOrder() As Byte
   Dim RdoSon As ADODB.Recordset
   On Error GoTo DiaErr1
   GetOldSalesOrder = False
   sSql = "SELECT SONUMBER FROM SohdTable WHERE SONUMBER=" & txtSon & " "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoSon)
   If bSqlRows Then
      With RdoSon
         GetOldSalesOrder = True
         ClearResultSet RdoSon
      End With
   End If
   Set RdoSon = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getoldsal"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   GetOldSalesOrder = False
   DoModuleErrors Me
   
End Function

'Private Sub AddSalesOrder()
'   Dim sNewDate As Variant
'   bGoodCust = GetCustomerData()
'   If bCutOff = 1 Then
'      MsgBox "This Customer's Credit Is On Hold.", _
'         vbInformation, Caption
'      bGoodCust = 0
'   End If
'   If Not bGoodCust Then Exit Sub
'   On Error GoTo DiaErr1
'   If Len(txtSon) < SO_NUM_SIZE Then
'      txtSon = Format(Abs(Val(txtSon)), SO_NUM_FORMAT)
'   End If
'   sNewDate = Format(ES_SYSDATE, "mm/dd/yy")
'   sSql = "INSERT SohdTable (SONUMBER,SOTYPE,SOCUST,SODATE," _
'          & "SOSALESMAN,SOSTNAME,SOSTADR,SODIVISION,SOREGION,SOSTERMS," _
'          & "SOVIA,SOFOB,SOARDISC,SODAYS,SONETDAYS,SOFREIGHTDAYS," _
'          & "SOTEXT,SOTAXEXEMPT,SOPO) " _
'          & "VALUES(" & Val(txtSon) & ",'" & cmbPre & "','" _
'          & sCust & "','" & sNewDate & "','" & sSlsMan & "','" _
'          & sStName & "','" & sStAdr & "','" & sDivision & "','" _
'          & sRegion & "','" & sSterms & "','" & sVia & "','" _
'          & sFob & "'," & cDiscount & "," & iDays & "," & iNetDays _
'          & "," & iFrtDays & ",'" & txtSon & "','" & sTaxExempt & "','" _
'          & Trim(txtCpo) & "')"
'
'   Debug.Print sSql
'
'   clsADOCon.ExecuteSql sSql 'rdExecDirect
'   If clsADOCon.RowsAffected Then
'      bSoAdded = 1
'      On Error Resume Next
'   '   MsgBox "Sales Order Added.", vbInformation, Caption
'      sSql = "UPDATE SohdTable SET SOCCONTACT='" & sContact & "'," _
'             & "SOCPHONE='" & sConPhone & "',SOCINTPHONE='" _
'             & sConIntPhone & "',SOCINTFAX='" & sConIntFax _
'             & "',SOCFAX='" & sConFax & "',SOCEXT=" & sConExt _
'             & " WHERE SONUMBER=" & Val(txtSon) & " "
'      clsADOCon.ExecuteSql sSql 'rdExecDirect
'      sSql = "UPDATE ComnTable SET COLASTSALESORDER='" & Trim(cmbPre) _
'             & Format(Trim(txtSon), SO_NUM_FORMAT) & "' WHERE COREF=1"
'      clsADOCon.ExecuteSql sSql 'rdExecDirect
'      optNew = vbChecked
'      SaleSLe02a.Show
'      SaleSLe02a.optNew = vbChecked
'
'      ' Save the last sales order revised so we can use it later (elsewhere)
'      SaveSetting "Esi2000", "EsiSale", "LastRevisedSO", txtSon
'
'      SaleSLe02a.cmbSon.SetFocus
'   Else
'      MsgBox "Couldn't Add Sales Order.", vbExclamation, Caption
'   End If
'   Exit Sub
'
'DiaErr1:
'   MsgBox Err.Description
'   If Left(Err.Description, 5) = "01000" Then
'      MsgBox "Sales Order Number " & txtSon & " Was Recently Used By  " & vbCrLf _
'         & "Another Process. Please Select The Next Number.", _
'         vbInformation, Caption
'      GetLastSalesOrder sOldSoNumber, sNewsonumber, True
'      sOldSoNumber = lblLst
'      lblNotice.Visible = False
'      tmr1.Enabled = True
'   Else
'      sProcName = "addsaleso"
'      CurrError.Number = Err.Number
'      CurrError.Description = Err.Description
'      DoModuleErrors Me
'   End If
'
'End Sub

Private Sub AddSalesOrder()
   Dim sNewDate As Variant
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
   
   'make sure no one is simultaneously creating the same SO number
   Dim so As Long
   so = CLng(txtSon)
   
   clsADOCon.BeginTrans
   Dim rs As ADODB.Recordset
   Do While True
      sSql = "select SONUMBER from SohdTable where SONUMBER = " & so
      If clsADOCon.GetDataSet(sSql, rs) = 0 Then Exit Do
      so = so + 1
      txtSon = so
   Loop
   
   sNewDate = Format(ES_SYSDATE, "mm/dd/yy")
   sSql = "INSERT SohdTable (SONUMBER,SOTYPE,SOCUST,SODATE," _
          & "SOSALESMAN,SOSTNAME,SOSTADR,SODIVISION,SOREGION,SOSTERMS," _
          & "SOVIA,SOFOB,SOARDISC,SODAYS,SONETDAYS,SOFREIGHTDAYS," _
          & "SOTAXEXEMPT,SOPO) " _
          & "VALUES(" & Val(txtSon) & ",'" & cmbPre & "','" _
          & sCust & "','" & sNewDate & "','" & sSlsMan & "','" _
          & sStName & "','" & sStAdr & "','" & sDivision & "','" _
          & sRegion & "','" & sSterms & "','" & sVia & "','" _
          & sFob & "'," & cDiscount & "," & iDays & "," & iNetDays _
          & "," & iFrtDays & ",'" & sTaxExempt & "','" _
          & Trim(txtCpo) & "')"
   
   clsADOCon.ExecuteSql sSql 'rdExecDirect
   If clsADOCon.RowsAffected Then
      bSoAdded = 1
      On Error Resume Next
   '   MsgBox "Sales Order Added.", vbInformation, Caption
      sSql = "UPDATE SohdTable SET SOCCONTACT='" & sContact & "'," _
             & "SOCPHONE='" & sConPhone & "',SOCINTPHONE='" _
             & sConIntPhone & "',SOCINTFAX='" & sConIntFax _
             & "',SOCFAX='" & sConFax & "',SOCEXT=" & sConExt _
             & " WHERE SONUMBER=" & Val(txtSon) & " "
      clsADOCon.ExecuteSql sSql 'rdExecDirect
      sSql = "UPDATE ComnTable SET COLASTSALESORDER='" & Trim(cmbPre) _
             & Format(Trim(txtSon), SO_NUM_FORMAT) & "' WHERE COREF=1"
      clsADOCon.ExecuteSql sSql 'rdExecDirect
      clsADOCon.CommitTrans
      
      ' go to revise SO transaction
      optNew = vbChecked
      SaleSLe02a.Show
      SaleSLe02a.optNew = vbChecked
          
      ' Save the last sales order revised so we can use it later (elsewhere)
      SaveSetting "Esi2000", "EsiSale", "LastRevisedSO", txtSon
          
      SaleSLe02a.cmbSon.SetFocus
   Else
      MsgBox "Couldn't Add Sales Order.", vbExclamation, Caption
   End If
   Exit Sub
   
DiaErr1:
   clsADOCon.RollbackTrans
   MsgBox Err.Description
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
          & "CUFRTDAYS,CUINTFAX,CUFAX,CUTAXEXEMPT,CUCUTOFF,CUSTCOUNTRY " _
          & "FROM CustTable WHERE CUREF='" & sCust & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCst)
   If bSqlRows Then
      With RdoCst
         bCutOff = !CUCUTOFF
         sStName = "" & Trim(!CUSTNAME)
         sStAdr = "" & Trim(!CUSTADR) & vbCrLf _
                  & "" & Trim(!CUSTCITY) & " " & Trim(!CUSTSTATE) _
                  & "  " & Trim(!CUSTZIP) _
                  & IIf(!CUSTCOUNTRY = "", "", vbCrLf + Trim(!CUSTCOUNTRY))
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
