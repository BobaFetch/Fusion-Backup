VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaARe10a
   BorderStyle = 3 'Fixed Dialog
   Caption = "Tax Codes"
   ClientHeight = 4410
   ClientLeft = 2115
   ClientTop = 1125
   ClientWidth = 4920
   ClipControls = 0 'False
   ControlBox = 0 'False
   ForeColor = &H80000007&
   LinkTopic = "Form1"
   MDIChild = -1 'True
   MinButton = 0 'False
   PaletteMode = 1 'UseZOrder
   ScaleHeight = 4410
   ScaleWidth = 4920
   ShowInTaskbar = 0 'False
   Begin VB.CommandButton cmdDeleteTax
      Caption = "&Delete"
      Height = 315
      Left = 3960
      TabIndex = 19
      TabStop = 0 'False
      ToolTipText = "Delete Current Account Number"
      Top = 720
      Width = 875
   End
   Begin VB.CommandButton cmdCan
      Cancel = -1 'True
      Caption = "Close"
      Height = 435
      Left = 3960
      TabIndex = 18
      TabStop = 0 'False
      Top = 120
      Width = 875
   End
   Begin VB.ComboBox cmbAccNum
      Height = 315
      Left = 1440
      TabIndex = 5
      Top = 2520
      Width = 1335
   End
   Begin VB.TextBox txtDesc
      Height = 735
      Left = 120
      TabIndex = 6
      Top = 3480
      Width = 4695
   End
   Begin VB.CheckBox chkBOMulti
      Caption = "___"
      ForeColor = &H8000000F&
      Height = 255
      Left = 1440
      TabIndex = 4
      Top = 2040
      Width = 735
   End
   Begin VB.Frame frmTaxType
      Height = 855
      Left = 2520
      TabIndex = 11
      Top = 360
      Width = 1215
      Begin VB.OptionButton optTypeBandO
         Caption = "B and O"
         Height = 255
         Left = 120
         TabIndex = 13
         Top = 480
         Width = 975
      End
      Begin VB.OptionButton optTypeSales
         Caption = "Sales"
         Height = 255
         Left = 120
         TabIndex = 12
         Top = 240
         Width = 855
      End
   End
   Begin VB.TextBox txtRate
      Height = 285
      Left = 1440
      TabIndex = 3
      Tag = "1"
      Top = 1560
      Width = 855
   End
   Begin VB.Frame Frame2
      Height = 30
      Left = 120
      TabIndex = 9
      Top = 1320
      Width = 4695
   End
   Begin VB.ComboBox cmbTaxCode
      Height = 315
      Left = 1080
      TabIndex = 2
      Top = 840
      Width = 1335
   End
   Begin VB.ComboBox cmbSte
      Height = 315
      Left = 1080
      TabIndex = 1
      Top = 480
      Width = 735
   End
   Begin Threed.SSRibbon cmdHlp
      Height = 225
      Left = 0
      TabIndex = 0
      ToolTipText = "Subject Help"
      Top = 0
      Width = 255
      _Version = 65536
      _ExtentX = 450
      _ExtentY = 397
      _StockProps = 65
      BackColor = 12632256
      GroupAllowAllUp = -1 'True
      Autosize = 2
      RoundedCorners = 0 'False
      BevelWidth = 0
      Outline = 0 'False
      PictureUp = "diaARe10.frx":0000
      PictureDn = "diaARe10.frx":0146
   End
   Begin ResizeLibCtl.ReSize ReSize1
      Left = 3840
      Top = 1920
      _Version = 196615
      _ExtentX = 741
      _ExtentY = 741
      _StockProps = 0
      Enabled = -1 'True
      FormMinWidth = 0
      FormMinHeight = 0
      FormDesignHeight = 4410
      FormDesignWidth = 4920
   End
   Begin VB.Label lblDsc
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      Height = 285
      Left = 1440
      TabIndex = 17
      Top = 2880
      Width = 3375
   End
   Begin VB.Label Label5
      Caption = "Comments"
      Height = 255
      Left = 120
      TabIndex = 16
      Top = 3240
      Width = 975
   End
   Begin VB.Label Label4
      Caption = "Account"
      Height = 255
      Left = 120
      TabIndex = 15
      Top = 2520
      Width = 1095
   End
   Begin VB.Label lblchk1
      Caption = "Report Multiple Activities?"
      Height = 495
      Left = 120
      TabIndex = 14
      Top = 1920
      Width = 1095
   End
   Begin VB.Label Label3
      Caption = "Tax Rate"
      Height = 255
      Left = 120
      TabIndex = 10
      Top = 1560
      Width = 735
   End
   Begin VB.Label Label2
      Caption = "Tax Code"
      Height = 255
      Left = 120
      TabIndex = 8
      Top = 840
      Width = 855
   End
   Begin VB.Label Label1
      Caption = "State Code"
      Height = 255
      Left = 120
      TabIndex = 7
      Top = 480
      Width = 975
   End
End
Attribute VB_Name = "diaARe10a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********************************************************************************
' diaARe08a - Tax Codes
'
' Created: (JH)
'
' Revisions:
'
'
'*********************************************************************************

Option Explicit
Dim bOnLoad As Byte
Dim bGoodCode As Byte
Dim rdoCode As rdoResultset

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd


Private Sub chkBOMulti_lostfocus()
   If bGoodCode Then
      On Error Resume Next
      rdoCode.Edit
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
      rdoCode.Edit
      rdoCode!TAXACCT = Compress(cmbAccNum)
      rdoCode.Update
      If Err > 0 Then ValidateEdit Me
   End If
End Sub

Private Sub cmbSte_Click()
   FillCodes
   GetCodes
End Sub

Private Sub cmbSte_LostFocus()
   FillCodes
   GetCodes
End Sub

Private Sub cmbTaxCode_LostFocus()
   cmbTaxCode = CheckLen(cmbTaxCode, 4)
   GetCodes
   
End Sub

Private Sub cmbTaxCode_Click()
   cmbTaxCode = CheckLen(cmbTaxCode, 4)
   GetCodes
End Sub

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdDeleteTax_Click()
   Dim bResponse As Byte
   
   bResponse = MsgBox("Are you sure you want to delete tax code " & cmbSte & " " & cmbTaxCode & "?", ES_YESQUESTION, Caption)
   
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
   rdoCode.Close
   On Error GoTo DiaErr1
   sSql = "DELETE TxcdTable " _
          & "WHERE TAXCODE = '" & Trim(cmbTaxCode) & "' AND TAXSTATE = '" & Trim(cmbSte) & "'"
   RdoCon.Execute sSql, rdExecDirect
   Manage
   bGoodCode = False
   cmbTaxCode.RemoveItem (0)
   cmbTaxCode = ""
   cmbTaxCode.SetFocus
   MouseCursor 0
   Exit Sub
   
   DiaErr1:
   sProcName = "Delete Tax Code"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub Form_Activate()
   
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      optTypeSales.Value = True
      FillStates Me
      FillAccounts
      
      bOnLoad = False
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   SetDiaPos Me
   FormatControls
   sCurrForm = Caption
   bOnLoad = True
   
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
   cmbTaxCode.Clear
   Dim CodeType As Integer
   Dim rdocmb As rdoResultset
   On Error GoTo DiaErr1
   
   If optTypeSales.Value = True Then
      CodeType = 1
   Else
      CodeType = 0
   End If
   
   If "" & Trim(cmbSte.Text) <> "" Then
      sSql = "SELECT TAXCODE FROM TxcdTable WHERE TAXTYPE = " & CodeType & " AND TAXSTATE = '" & cmbSte & "'"
      bSqlRows = GetDataSet(rdocmb, ES_KEYSET)
      If bSqlRows Then
         With rdocmb
            cmbTaxCode = "" & Trim(!TAXCODE)
            Do Until .EOF
               cmbTaxCode.AddItem "" & Trim(!TAXCODE)
               .MoveNext
            Loop
            .Cancel
         End With
      End If
      Set rdocmb = Nothing
   End If
   Exit Sub
   
   DiaErr1:
   sProcName = "FillCodes"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub optTypeBandO_Click()
   FillCodes
   cmbAccNum.Enabled = False
   chkBOMulti.Enabled = True
   GetCodes
End Sub

Private Sub optTypeSales_Click()
   FillCodes
   cmbAccNum.Enabled = True
   chkBOMulti.Enabled = False
   GetCodes
End Sub

Private Sub GetCodes()
   Dim bResponse As Byte
   Dim bCodeType As Byte
   
   If optTypeSales.Value = True Then
      bCodeType = 1
   Else
      bCodeType = 0
   End If
   
   Manage
   
   If "" & Trim(cmbSte) <> "" And "" & Trim(cmbTaxCode) <> "" Then
      sSql = "SELECT TAXDESC, TAXACCT, TAXCOUNTY, TAXRATE, TAXTYPE, TAXMULTIPLE, TAXREF FROM TxcdTable WHERE TAXCODE = '" & cmbTaxCode & "' AND TAXSTATE = '" & cmbSte & "'"
      bSqlRows = GetDataSet(rdoCode, ES_KEYSET)
      If bSqlRows Then
         With rdoCode
            If !TAXTYPE <> 0 And bCodeType <> 0 Then
               txtRate = !TAXRATE
               cmbAccNum = "" & Trim(!TAXACCT)
               txtDesc = "" & Trim(!TAXDESC)
               chkBOMulti.Value = vbUnchecked
               bGoodCode = True
            ElseIf !TAXTYPE = 0 And bCodeType = 0 Then
               txtRate = !TAXRATE
               txtDesc = "" & Trim(!TAXDESC)
               chkBOMulti.Value = !TAXMULTIPLE
               bGoodCode = True
            Else
               If bCodeType <> 0 Then
                  MsgBox "This is a B & O Tax Code!", vbExclamation, Caption
                  cmbTaxCode.SetFocus
                  Manage
                  Exit Sub
               Else
                  MsgBox "This is a Sales Tax Code!", vbExclamation, Caption
                  cmbTaxCode.SetFocus
                  Manage
                  Exit Sub
               End If
            End If
         End With
      Else
         bResponse = MsgBox(cmbSte & " " & cmbTaxCode & " Wasn't Found. Add It?", ES_YESQUESTION, Caption)
         If bResponse <> vbYes Then
            bGoodCode = False
            MouseCursor 0
            CancelTrans
            cmbTaxCode = ""
            On Error Resume Next
            cmdCan.SetFocus
            Exit Sub
         End If
         MouseCursor 13
         On Error Resume Next
         rdoCode.Close
         On Error GoTo DiaErr1
         Dim TaxRefference As String
         TaxRefference = Compress(cmbSte & cmbTaxCode)
         sSql = "INSERT TxcdTable (TAXCODE, TAXSTATE, TAXTYPE, TAXREF) " _
                & "VALUES('" & Trim(cmbTaxCode) & "','" _
                & Trim(cmbSte) & "', " & bCodeType & ", '" & TaxRefference & "')"
         RdoCon.Execute sSql, rdExecDirect
         Manage
         cmbTaxCode.AddItem "" & Trim(cmbTaxCode)
         bGoodCode = True
         cmbTaxCode.SetFocus
         MouseCursor 0
         Exit Sub
      End If
   End If
   Exit Sub
   
   DiaErr1:
   CurrError.Description = Err.Description
   Resume DiaErr2
   DiaErr2:
   bGoodCode = False
   Manage
   cmbSte = ""
   MouseCursor 0
   MsgBox CurrError.Description & vbCrLf & "Couldn't Add Tax Code.", vbExclamation, Caption
   
End Sub

Private Sub Manage()
   txtRate.Text = "0.00"
   cmbAccNum.Text = ""
   chkBOMulti.Value = False
   txtDesc.Text = ""
   lblDsc.Caption = ""
End Sub

Public Sub FillAccounts()
   ' Fill account combo
   ' Need to add account descriptions
   Dim rdoAct As rdoResultset
   
   sSql = "Qry_FillLowAccounts"
   bSqlRows = GetDataSet(rdoAct, ES_FORWARD)
   
   If bSqlRows Then
      With rdoAct
         Do Until .EOF
            AddComboStr cmbAccNum.hWnd, "" & Trim(!GLACCTNO)
            .MoveNext
         Loop
      End With
      cmbAccNum.ListIndex = 0
   End If
   Set rdoAct = Nothing
End Sub

Private Sub txtDesc_LostFocus()
   txtDesc = StrCase(txtDesc)
   txtDesc = CheckLen(txtDesc, 40)
   If bGoodCode Then
      On Error Resume Next
      rdoCode.Edit
      rdoCode!TAXDESC = "" & txtDesc
      rdoCode.Update
      If Err > 0 Then ValidateEdit Me
   End If
End Sub

Private Sub txtRate_LostFocus()
   txtRate = Format(txtRate, "0.000")
   If bGoodCode Then
      On Error Resume Next
      rdoCode.Edit
      rdoCode!TAXRATE = Val(txtRate)
      rdoCode.Update
      If Err > 0 Then ValidateEdit Me
   End If
   
End Sub


Function GetAccount(sCmbAcct As String) As Byte
   Dim rdoGlm As rdoResultset
   On Error GoTo DiaErr1
   
   sCmbAcct = Compress(sCmbAcct)
   
   If sCmbAcct = "" Then Exit Function
   
   sSql = "SELECT GLACCTREF,GLACCTNO,GLDESCR FROM GlacTable " _
          & "WHERE GLACCTREF='" & sCmbAcct & "' AND GLINACTIVE=0"
   bSqlRows = GetDataSet(rdoGlm, ES_FORWARD)
   
   If bSqlRows Then
      cmbAccNum = "" & Trim(rdoGlm!GLACCTNO)
      lblDsc = "" & Trim(rdoGlm!GLDESCR)
      GetAccount = 1
   Else
      lblDsc = "*** Account Wasn't Found Or Inactive ***"
      cmbAccNum.SetFocus
   End If
   Set rdoGlm = Nothing
   Exit Function
   
   DiaErr1:
   sProcName = "GetAccount"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function
