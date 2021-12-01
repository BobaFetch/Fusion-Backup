VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaARe10a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Revise Tax Codes"
   ClientHeight    =   4395
   ClientLeft      =   5535
   ClientTop       =   5190
   ClientWidth     =   6330
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H80000007&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4395
   ScaleWidth      =   6330
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdDeleteTax 
      Caption         =   "&Delete"
      Height          =   315
      Left            =   5280
      TabIndex        =   10
      TabStop         =   0   'False
      ToolTipText     =   "Delete Current Account Number"
      Top             =   720
      Width           =   875
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5280
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   240
      Width           =   875
   End
   Begin VB.ComboBox cmbAccNum 
      Height          =   315
      Left            =   1440
      TabIndex        =   7
      Top             =   2520
      Width           =   1575
   End
   Begin VB.TextBox txtDesc 
      Height          =   735
      Left            =   120
      TabIndex        =   8
      Top             =   3480
      Width           =   4695
   End
   Begin VB.CheckBox chkBOMulti 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   1440
      TabIndex        =   6
      Top             =   2040
      Value           =   2  'Grayed
      Width           =   735
   End
   Begin VB.Frame frmTaxType 
      Height          =   1095
      Left            =   3840
      TabIndex        =   16
      Top             =   120
      Width           =   1215
      Begin VB.OptionButton optTypeCountry 
         Caption         =   "Country"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   975
      End
      Begin VB.OptionButton optTypeBandO 
         Caption         =   "B and O"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   975
      End
      Begin VB.OptionButton optTypeSales 
         Caption         =   "Sales"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.TextBox txtRate 
      Height          =   285
      Left            =   1440
      TabIndex        =   5
      Tag             =   "1"
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Height          =   30
      Left            =   0
      TabIndex        =   14
      Top             =   1320
      Width           =   4935
   End
   Begin VB.ComboBox cmbTaxCode 
      Height          =   315
      ItemData        =   "diaARe10a.frx":0000
      Left            =   1080
      List            =   "diaARe10a.frx":0002
      TabIndex        =   1
      Top             =   840
      Width           =   2475
   End
   Begin VB.ComboBox cmbSte 
      Height          =   315
      ItemData        =   "diaARe10a.frx":0004
      Left            =   1080
      List            =   "diaARe10a.frx":0006
      TabIndex        =   0
      Top             =   480
      Width           =   2475
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   11
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
      PictureUp       =   "diaARe10a.frx":0008
      PictureDn       =   "diaARe10a.frx":014E
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   4320
      Top             =   1440
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4395
      FormDesignWidth =   6330
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      Height          =   255
      Index           =   3
      Left            =   2640
      TabIndex        =   21
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1440
      TabIndex        =   20
      Top             =   2880
      Width           =   3000
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Comments"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   19
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Account"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   18
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Report Multiple Activities?"
      Height          =   495
      Index           =   4
      Left            =   120
      TabIndex        =   17
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Rate"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   15
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Code"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   13
      Top             =   840
      Width           =   855
   End
   Begin VB.Label lblState 
      BackStyle       =   0  'Transparent
      Caption         =   "State Code"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   480
      Width           =   975
   End
End
Attribute VB_Name = "diaARe10a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of                     ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***

'See the UpdateTables prodecure for database revisions

'**********************************************************************************
' diaARe10a - Add Revise Tax Codes
'
' Created: 09/09/02 (JH)
'
' Revisions:
'   10/15/02 (nth) Added to esifina.
'   01/29/03 (nth) Changed tax rate format from "0.000" to "0.#####".
'
'**********************************************************************************

Option Explicit

Dim bOnLoad As Byte
Dim bGoodCode As Byte
Dim rdoCode As ADODB.Recordset
Dim bCodeType As Byte

Private inFillCodes As Boolean
Private inGetCodes As Boolean

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'**********************************************************************************

Private Sub chkBOMulti_lostfocus()
   If bGoodCode Then
      On Error Resume Next
      rdoCode!TAXMULTIPLE = chkBOMulti.Value
      rdoCode.Update
      If Err > 0 Then ValidateEdit Me
   End If
End Sub

Private Sub cmbAccNum_Click()
   GetAccount cmbAccNum
End Sub

Private Sub cmbAccNum_LostFocus()
   cmbAccNum = CheckLen(cmbAccNum, 12)
   Dim bCheck As Byte
   bCheck = GetAccount(cmbAccNum)
   If bGoodCode And Trim(cmbAccNum) <> "" And bCheck Then
      On Error Resume Next
      rdoCode!TAXACCT = Compress(cmbAccNum)
      rdoCode.Update
      If Err > 0 Then ValidateEdit Me
   End If
End Sub

Private Sub cmbSte_Click()
   Debug.Print "BO=" & optTypeBandO.Value & " Sales=" & optTypeSales.Value & " Country=" & optTypeCountry.Value
   FillCodes
   GetCodes
End Sub

Private Sub cmbSte_GotFocus()
   Debug.Print "BO=" & optTypeBandO.Value & " Sales=" & optTypeSales.Value & " Country=" & optTypeCountry.Value
End Sub

Private Sub cmbSte_LostFocus()
   'Debug.Print "BO=" & optTypeBandO.Value & " Sales=" & optTypeSales.Value & " Country=" & optTypeCountry.Value
   'FillCodes
   'GetCodes
End Sub

Private Sub cmbTaxCode_LostFocus()
   cmbTaxCode = CheckLen(cmbTaxCode, 8)
   GetCodes
End Sub

Private Sub cmbTaxCode_Click()
   cmbTaxCode = CheckLen(cmbTaxCode, 8)
   GetCodes
End Sub

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdDeleteTax_Click()
   DeleteCode
End Sub

Private Sub Form_Activate()
   
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      
      FillAccounts
      FillStates Me
      
      bOnLoad = False
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   sCurrForm = Caption
   bOnLoad = True
   txtRate = "0.00000"
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set rdoCode = Nothing
   Set diaARe10a = Nothing
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Public Sub FillCodes()
   
   '    If inFillCodes Then
   '        Exit Sub
   '    Else
   '        inFillCodes = True
   '    End If
   
   Dim RdoCmb As ADODB.Recordset
   
   On Error GoTo DiaErr1
   cmbTaxCode.Clear
   
   If optTypeBandO.Value = True Then
      bCodeType = 0 'b & o tax
      If cmbSte.ListCount < 50 Then
         cmbSte.Clear
         FillStates Me
         lblState.Caption = "State"
      End If
   ElseIf optTypeSales.Value = True Then
      bCodeType = 1 'sales tax
      If cmbSte.ListCount < 50 Then
         cmbSte.Clear
         FillStates Me
         lblState.Caption = "State"
      End If
   Else
      'cmbSte.Clear
      bCodeType = 2 'country tax
      If cmbSte.ListCount = 0 Or cmbSte.ListCount >= 50 Then
         FillCountries
         lblState.Caption = "Country"
      End If
   End If
   
   sSql = "SELECT TAXCODE, TAXSTATE FROM TxcdTable WHERE TAXTYPE = " & bCodeType
   'If bCodeType <> 2 Then
   sSql = sSql & " AND TAXSTATE = '" & cmbSte & "'"
   'End If
   
   Debug.Print sSql
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCmb)
   If bSqlRows Then
      With RdoCmb
         While Not .EOF
            AddComboStr cmbTaxCode.hwnd, "" & Trim(!taxCode)
            'cmbSte.AddItem "" & Trim(!TAXSTATE)
            .MoveNext
         Wend
      End With
      cmbTaxCode.ListIndex = 0
   End If
   Set RdoCmb = Nothing
   inFillCodes = False
   
   Exit Sub
   
DiaErr1:
   inFillCodes = False
   sProcName = "FillCodes"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub FillCountries()
   Dim RdoCmb As ADODB.Recordset
   cmbSte.Clear
   sSql = "SELECT DISTINCT TAXSTATE FROM TxcdTable WHERE TAXTYPE = 2 ORDER BY TAXSTATE"
   Debug.Print sSql
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCmb)
   If bSqlRows Then
      'If cmbSte.ListCount <> RdoCmb.RowCount Then
      'cmbSte.Clear
      With RdoCmb
         While Not .EOF
            'AddComboStr cmbTaxCode.hWnd, "" & Trim(!TAXSTATE)
            cmbSte.AddItem "" & Trim(!taxState)
            .MoveNext
         Wend
      End With
      'If cmbSte.ListCount > 0 Then
      '    cmbSte.ListIndex = 0
      'End If
      'End If
   End If
   Set RdoCmb = Nothing
   
End Sub

Private Sub optTypeBandO_Click()
   cmbTaxCode = ""
   FillCodes
   cmbAccNum.enabled = False
   chkBOMulti.enabled = True
   GetCodes
End Sub

Private Sub optTypeBandO_LostFocus()
   Debug.Print "BO=" & optTypeBandO.Value & " Sales=" & optTypeSales.Value & " Country=" & optTypeCountry.Value
End Sub

Private Sub optTypeCountry_Click()
   cmbTaxCode = ""
   FillCodes
   cmbAccNum.enabled = True
   chkBOMulti.enabled = False
   GetCodes
End Sub

Private Sub optTypeCountry_LostFocus()
   Debug.Print "BO=" & optTypeBandO.Value & " Sales=" & optTypeSales.Value & " Country=" & optTypeCountry.Value
End Sub

Private Sub optTypeSales_Click()
   cmbTaxCode = ""
   FillCodes
   cmbAccNum.enabled = True
   chkBOMulti.enabled = False
   GetCodes
End Sub

Private Sub GetCodes()
   
   '    If inGetCodes Then
   '        Exit Sub
   '    Else
   '        inGetCodes = True
   '    End If
   
   Dim RdoCmb As ADODB.Recordset
   
   
   Dim bResponse As Byte
   
   If optTypeSales.Value = True Then
      bCodeType = 1
   ElseIf optTypeCountry.Value = True Then
      bCodeType = 2
   Else
      bCodeType = 0
   End If
   
   Manage
   On Error GoTo DiaErr1
   If Trim(cmbSte) <> "" And Trim(cmbTaxCode) <> "" Then
      sSql = "SELECT TAXDESC, TAXACCT, TAXCOUNTY, TAXRATE, TAXTYPE, TAXMULTIPLE," _
             & "TAXREF FROM TxcdTable WHERE TAXCODE = '" & Trim(cmbTaxCode) & "'" _
             & " AND TAXSTATE = '" & Trim(cmbSte) & "'"
      '        If optTypeCountry.Value Then
      '            sSql = sSql & " AND TAXCOUNTRY = '" & Trim(cmbSte) & "'"
      '        Else
      '            sSql = sSql & " AND TAXSTATE = '" & Trim(cmbSte) & "'"
      '        End If
      bSqlRows = clsADOCon.GetDataSet(sSql, rdoCode, ES_KEYSET)
      
      If bSqlRows Then
         With rdoCode
            If !TAXTYPE <> 0 And bCodeType <> 0 Then
               txtRate = Format(!TAXRATE, "0.#####")
               cmbAccNum = "" & Trim(!TAXACCT)
               GetAccount "" & Trim(!TAXACCT)
               txtDesc = "" & Trim(!TAXDESC)
               chkBOMulti.Value = vbUnchecked
               chkBOMulti.enabled = False
               bGoodCode = True
            ElseIf !TAXTYPE = 0 And bCodeType = 0 Then
               txtRate = Format(!TAXRATE, "0.#####")
               txtDesc = "" & Trim(!TAXDESC)
               chkBOMulti.Value = !TAXMULTIPLE
               chkBOMulti.enabled = True
               bGoodCode = True
            Else
               If bCodeType <> 0 Then
                  MsgBox "This is a B & O Tax Code.", vbInformation, Caption
                  cmbTaxCode.SetFocus
                  Manage
                  inGetCodes = False
                  Exit Sub
               Else
                  MsgBox "This is a Sales Tax Code.", vbInformation, Caption
                  cmbTaxCode.SetFocus
                  Manage
                  Exit Sub
               End If
            End If
         End With
      Else
         'Set rdoCode = Nothing
         
         'adding code -- check for 12 character limit
         Dim sCode As String
         sCode = Trim(Compress(cmbSte & cmbTaxCode))
         If Len(sCode) > 12 Then
            MsgBox "State/Country plus tax code (" & sCode & ") is limited to 12 characters", vbExclamation, Caption
            inGetCodes = False
            Exit Sub
         End If
         
         bResponse = MsgBox(cmbSte & " " & cmbTaxCode _
                     & " Wasn't Found. Add It ?", ES_YESQUESTION, Caption)
         
         If bResponse <> vbYes Then
            bGoodCode = False
            MouseCursor 0
            CancelTrans
            cmbTaxCode = ""
            On Error Resume Next
            cmdCan.SetFocus
            inGetCodes = False
            Exit Sub
         End If
         AddCode
         txtRate.SetFocus
         inGetCodes = False
         Exit Sub
      End If
   End If
   inGetCodes = False
   Exit Sub
   
DiaErr1:
   inGetCodes = False
   CurrError.Description = Err.Description
   Resume DiaErr2
DiaErr2:
   bGoodCode = False
   Manage
   cmbSte = ""
   MouseCursor 0
   MsgBox CurrError.Description & vbCrLf & "Couldn't Add Tax Code.", _
      vbExclamation, Caption
End Sub

Private Sub Manage()
   txtRate.Text = "0.000"
   cmbAccNum.Text = ""
   chkBOMulti.Value = vbUnchecked
   If bCodeType = 1 Then
      chkBOMulti.enabled = False
   Else
      chkBOMulti.enabled = True
   End If
   txtDesc.Text = ""
   lblDsc.Caption = ""
End Sub

Public Sub FillAccounts()
   Dim rdoAct As ADODB.Recordset
   
   On Error GoTo DiaErr1
   sSql = "Qry_FillLowAccounts"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoAct, ES_FORWARD)
   
   If bSqlRows Then
      With rdoAct
         Do Until .EOF
            AddComboStr cmbAccNum.hwnd, "" & Trim(!GLACCTNO)
            .MoveNext
         Loop
      End With
      cmbAccNum = ""
   End If
   Set rdoAct = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "fillaccounts"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub optTypeSales_LostFocus()
   Debug.Print "BO=" & optTypeBandO.Value & " Sales=" & optTypeSales.Value & " Country=" & optTypeCountry.Value
End Sub

Private Sub txtDesc_LostFocus()
   txtDesc = StrCase(txtDesc)
   txtDesc = CheckLen(txtDesc, 255)
   If bGoodCode Then
      On Error Resume Next
      rdoCode!TAXDESC = "" & txtDesc
      rdoCode.Update
      If Err > 0 Then ValidateEdit Me
   End If
End Sub

Private Sub txtRate_GotFocus()
   SelectFormat Me
End Sub

Private Sub txtRate_LostFocus()
   txtRate = Format(txtRate, "0.00##")
   If bGoodCode Then
      On Error Resume Next
      rdoCode!TAXRATE = Val(txtRate)
      rdoCode.Update
      If Err > 0 Then ValidateEdit Me
   End If
End Sub


Function GetAccount(sCmbAcct As String) As Byte
   Dim RdoGlm As ADODB.Recordset
   On Error GoTo DiaErr1
   
   sCmbAcct = Compress(sCmbAcct)
   
   If sCmbAcct = "" Then Exit Function
   
   sSql = "SELECT GLACCTREF,GLACCTNO,GLDESCR FROM GlacTable " _
          & "WHERE GLACCTREF='" & sCmbAcct & "' AND GLINACTIVE=0"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoGlm, ES_FORWARD)
   
   If bSqlRows Then
      lblDsc = "" & Trim(RdoGlm!GLDESCR)
      GetAccount = 1
   Else
      lblDsc = "*** Account Wasn't Found Or Inactive ***"
      cmbAccNum.SetFocus
   End If
   Set RdoGlm = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "GetAccount"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Function

Public Sub AddCode()
   Dim TaxReference As String
   Dim taxCode As String
   
   taxCode = Trim(cmbTaxCode)
   TaxReference = Trim(Compress(cmbSte & cmbTaxCode))
   If Len(TaxReference) > 12 Then
      MsgBox "State/Country plus tax code (" & TaxReference & ") is limited to 12 characters", vbExclamation, Caption
      Exit Sub
   End If
   
   MouseCursor 13
   On Error GoTo DiaErr1
   
   Dim taxState As String
   taxState = Trim(cmbSte)
   
   sSql = "INSERT INTO TxcdTable (TAXCODE, TAXSTATE, TAXTYPE, TAXREF) " _
          & "VALUES('" & taxCode & "','" _
          & taxState & "', " & bCodeType & ", '" & TaxReference & "')"
   
   Debug.Print sSql
   clsADOCon.ExecuteSQL sSql
   
   Manage
   
   'if state/country not in combobox, add it
   Dim i As Integer
   Dim found As Boolean
   found = False
   For i = 0 To cmbSte.ListCount - 1
      If cmbSte.List(i) = taxState Then
         found = True
         Exit For
      End If
   Next
   If Not found Then
      cmbSte.AddItem taxState
      cmbSte.ListIndex = cmbSte.ListCount - 1
   End If
   
   'if taxcode not in combobox, add it
   found = False
   For i = 0 To cmbTaxCode.ListCount - 1
      If cmbTaxCode.List(i) = taxCode Then
         found = True
         Exit For
      End If
   Next
   If Not found Then
      cmbTaxCode.AddItem taxCode
      cmbTaxCode.ListIndex = cmbTaxCode.ListCount - 1
   End If
   
   FillCodes
   cmbTaxCode = taxCode
   GetCodes
   bGoodCode = True
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "addcode"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub


Public Sub DeleteCode()
   Dim bResponse As Byte
   Dim sCode As String
   
   bResponse = MsgBox("Delete Tax Code " & cmbSte & " " _
               & cmbTaxCode & " ?", ES_YESQUESTION, Caption)
   
   If bResponse <> vbYes Then
      bGoodCode = True
      MouseCursor 0
      CancelTrans
      On Error Resume Next
      cmdCan.SetFocus
      Exit Sub
   End If
   
   MouseCursor 13
   On Error Resume Next
   Set rdoCode = Nothing
   
   On Error GoTo DiaErr1
   sSql = "DELETE TxcdTable " _
          & "WHERE TAXCODE = '" & Trim(cmbTaxCode) & "' AND TAXSTATE = '" & Trim(cmbSte) & "'"
   clsADOCon.ExecuteSQL sSql
   Manage
   bGoodCode = False
   cmbTaxCode.RemoveItem (0)
   cmbTaxCode = ""
   cmbTaxCode.SetFocus
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "DeleteCode"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub
