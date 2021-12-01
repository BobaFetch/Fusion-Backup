VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaGLe03a
   BorderStyle = 3 'Fixed Dialog
   Caption = "Financial Statement Structure"
   ClientHeight = 5715
   ClientLeft = 45
   ClientTop = 330
   ClientWidth = 6825
   ClipControls = 0 'False
   ControlBox = 0 'False
   LinkTopic = "Form1"
   MaxButton = 0 'False
   MDIChild = -1 'True
   MinButton = 0 'False
   ScaleHeight = 5715
   ScaleWidth = 6825
   ShowInTaskbar = 0 'False
   Begin VB.CommandButton cmdUpd
      Caption = "&Update"
      Height = 315
      Left = 5880
      TabIndex = 40
      ToolTipText = "Update Structure And Associated Entries"
      Top = 600
      Width = 875
   End
   Begin VB.TextBox txtFdtx
      Height = 285
      Left = 720
      TabIndex = 10
      Tag = "3"
      Top = 4920
      Width = 1675
   End
   Begin VB.TextBox txtOexp
      Height = 285
      Left = 720
      TabIndex = 9
      Tag = "3"
      Top = 4320
      Width = 1675
   End
   Begin VB.TextBox txtOinc
      Height = 285
      Left = 720
      TabIndex = 8
      Tag = "3"
      Top = 3960
      Width = 1675
   End
   Begin VB.TextBox txtExpn
      Height = 285
      Left = 720
      TabIndex = 7
      Tag = "3"
      Top = 3240
      Width = 1675
   End
   Begin VB.TextBox txtCogs
      Height = 285
      Left = 720
      TabIndex = 6
      Tag = "3"
      Top = 2520
      Width = 1675
   End
   Begin VB.TextBox txtIncm
      Height = 285
      Left = 720
      TabIndex = 5
      Tag = "3"
      Top = 2160
      Width = 1675
   End
   Begin VB.TextBox txtEqty
      Height = 285
      Left = 720
      TabIndex = 4
      Tag = "3"
      Top = 1560
      Width = 1675
   End
   Begin VB.TextBox txtLiab
      Height = 285
      Left = 720
      TabIndex = 3
      Tag = "3"
      Top = 1080
      Width = 1675
   End
   Begin VB.TextBox txtAsst
      Height = 285
      Left = 720
      TabIndex = 0
      Tag = "3"
      Top = 720
      Width = 1675
   End
   Begin VB.CommandButton cmdCan
      Cancel = -1 'True
      Caption = "Close"
      Height = 435
      Left = 5880
      TabIndex = 1
      TabStop = 0 'False
      Top = 120
      Width = 875
   End
   Begin Threed.SSRibbon cmdHlp
      Height = 225
      Left = 0
      TabIndex = 2
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
      PictureUp = "diaGlmst.frx":0000
      PictureDn = "diaGlmst.frx":0146
   End
   Begin ResizeLibCtl.ReSize ReSize1
      Left = 6480
      Top = 5280
      _Version = 196615
      _ExtentX = 741
      _ExtentY = 741
      _StockProps = 0
      Enabled = -1 'True
      FormMinWidth = 0
      FormMinHeight = 0
      FormDesignHeight = 5715
      FormDesignWidth = 6825
   End
   Begin VB.Label Z1
      BackStyle = 0 'Transparent
      Caption = "(Blank If None)"
      Height = 255
      Index = 23
      Left = 5400
      TabIndex = 44
      Top = 4920
      Width = 1575
   End
   Begin VB.Label Z1
      BackStyle = 0 'Transparent
      Caption = "(Blank If None)"
      Height = 255
      Index = 22
      Left = 5400
      TabIndex = 43
      Top = 4320
      Width = 1575
   End
   Begin VB.Label Z1
      BackStyle = 0 'Transparent
      Caption = "(Blank If None)"
      Height = 255
      Index = 21
      Left = 5400
      TabIndex = 42
      Top = 3960
      Width = 1575
   End
   Begin VB.Label Z1
      BackStyle = 0 'Transparent
      Caption = "(Blank If None)"
      Height = 255
      Index = 19
      Left = 5400
      TabIndex = 41
      Top = 2520
      Width = 1575
   End
   Begin VB.Label Z1
      Alignment = 2 'Center
      Caption = "-"
      Height = 255
      Index = 20
      Left = 240
      TabIndex = 39
      Top = 4920
      Width = 375
   End
   Begin VB.Label Z1
      Alignment = 2 'Center
      Caption = "-"
      Height = 255
      Index = 18
      Left = 240
      TabIndex = 38
      Top = 4320
      Width = 375
   End
   Begin VB.Label Z1
      Alignment = 2 'Center
      Caption = "+"
      Height = 255
      Index = 17
      Left = 240
      TabIndex = 37
      Top = 3960
      Width = 375
   End
   Begin VB.Label Z1
      Alignment = 2 'Center
      Caption = "="
      Height = 255
      Index = 16
      Left = 240
      TabIndex = 36
      Top = 5280
      Width = 375
   End
   Begin VB.Label Z1
      BackStyle = 0 'Transparent
      Caption = "Net Operating Profit (With Cost of Goods and Other Expense)"
      Height = 210
      Index = 15
      Left = 720
      TabIndex = 35
      Top = 5280
      Width = 5415
   End
   Begin VB.Label Z1
      Alignment = 2 'Center
      Caption = "="
      Height = 255
      Index = 14
      Left = 240
      TabIndex = 34
      Top = 4680
      Width = 375
   End
   Begin VB.Label Z1
      BackStyle = 0 'Transparent
      Caption = "Net Profit (Pretax Net)"
      Height = 210
      Index = 13
      Left = 720
      TabIndex = 33
      Top = 4680
      Width = 5415
   End
   Begin VB.Label Z1
      Alignment = 2 'Center
      Caption = "="
      Height = 255
      Index = 12
      Left = 240
      TabIndex = 32
      Top = 3650
      Width = 375
   End
   Begin VB.Label Z1
      Alignment = 2 'Center
      BackStyle = 0 'Transparent
      Caption = "-------------------------------------"
      Height = 210
      Index = 11
      Left = 720
      TabIndex = 31
      Top = 3480
      Width = 1695
   End
   Begin VB.Label Z1
      BackStyle = 0 'Transparent
      Caption = "Net Operating Profit (With Cost of Goods and Other Expense)"
      Height = 210
      Index = 10
      Left = 720
      TabIndex = 30
      Top = 3650
      Width = 5415
   End
   Begin VB.Label Z1
      Alignment = 2 'Center
      Caption = "="
      Height = 255
      Index = 9
      Left = 240
      TabIndex = 29
      Top = 2900
      Width = 375
   End
   Begin VB.Label Z1
      BackStyle = 0 'Transparent
      Caption = "Gross Profit (If Cost of Goods Sold Included"
      Height = 210
      Index = 8
      Left = 720
      TabIndex = 28
      Top = 2900
      Width = 4215
   End
   Begin VB.Label Z1
      BackStyle = 0 'Transparent
      Caption = "Enter/Revise Financial Statement Structure:"
      Height = 255
      Index = 7
      Left = 240
      TabIndex = 27
      Top = 1920
      Width = 3855
   End
   Begin VB.Label Z1
      BackStyle = 0 'Transparent
      Caption = "Enter/Revise Balance Sheet Structure:"
      Height = 255
      Index = 6
      Left = 240
      TabIndex = 26
      Top = 360
      Width = 3855
   End
   Begin VB.Label Z1
      Alignment = 2 'Center
      Caption = "-"
      Height = 255
      Index = 5
      Left = 240
      TabIndex = 25
      Top = 3240
      Width = 375
   End
   Begin VB.Label Z1
      Alignment = 2 'Center
      Caption = "="
      Height = 255
      Index = 4
      Left = 240
      TabIndex = 24
      Top = 1560
      Width = 375
   End
   Begin VB.Label Z1
      Alignment = 2 'Center
      BackStyle = 0 'Transparent
      Caption = "-------------------------------------"
      Height = 210
      Index = 3
      Left = 720
      TabIndex = 23
      Top = 2760
      Width = 1695
   End
   Begin VB.Label Z1
      Alignment = 2 'Center
      Caption = "-"
      Height = 255
      Index = 2
      Left = 240
      TabIndex = 22
      Top = 2520
      Width = 375
   End
   Begin VB.Label Z1
      Alignment = 2 'Center
      BackStyle = 0 'Transparent
      Caption = "-------------------------------------"
      Height = 150
      Index = 1
      Left = 720
      TabIndex = 21
      Top = 1320
      Width = 1695
   End
   Begin VB.Label Z1
      Alignment = 2 'Center
      Caption = "-"
      Height = 255
      Index = 0
      Left = 240
      TabIndex = 20
      Top = 1080
      Width = 375
   End
   Begin VB.Label lblFdtx
      BackStyle = 0 'Transparent
      Caption = "Federal Income Tax Master Account"
      Height = 255
      Left = 2640
      TabIndex = 19
      Top = 4920
      Width = 2895
   End
   Begin VB.Label lblOexp
      BackStyle = 0 'Transparent
      Caption = "Other Expense Master Account"
      Height = 255
      Left = 2640
      TabIndex = 18
      Top = 4320
      Width = 2895
   End
   Begin VB.Label lblOinc
      BackStyle = 0 'Transparent
      Caption = "Other Income Master Account"
      Height = 255
      Left = 2640
      TabIndex = 17
      Top = 3960
      Width = 2895
   End
   Begin VB.Label lblExpn
      BackStyle = 0 'Transparent
      Caption = "Expense Master Account"
      Height = 255
      Left = 2640
      TabIndex = 16
      Top = 3240
      Width = 2895
   End
   Begin VB.Label lblCogs
      BackStyle = 0 'Transparent
      Caption = "Cost of Goods Sold Master Account"
      Height = 255
      Left = 2640
      TabIndex = 15
      Top = 2520
      Width = 2895
   End
   Begin VB.Label lblIncm
      BackStyle = 0 'Transparent
      Caption = "Income Master Account"
      Height = 255
      Left = 2640
      TabIndex = 14
      Top = 2160
      Width = 2895
   End
   Begin VB.Label lblEqty
      BackStyle = 0 'Transparent
      Caption = "Equity Master Account"
      Height = 255
      Left = 2640
      TabIndex = 13
      Top = 1560
      Width = 2895
   End
   Begin VB.Label lblLiab
      BackStyle = 0 'Transparent
      Caption = "Liability Master Account"
      Height = 255
      Left = 2640
      TabIndex = 12
      Top = 1080
      Width = 2655
   End
   Begin VB.Label lblAsst
      BackStyle = 0 'Transparent
      Caption = "Asset Master Account"
      Height = 255
      Left = 2640
      TabIndex = 11
      Top = 720
      Width = 2655
   End
End
Attribute VB_Name = "diaGLe03a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001, ES/2002) is the property of                     ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***

'See the UpdateTables prodecure for database revisions

Option Explicit

'*********************************************************************************
'   diaGLe03a - Create/Revise finacial statment structure
'
'   Notes:
'
'   Created: (cjs)
'   Revisions:
'       02/01/02 (nth) Determine if the statment structure has never
'                      been setup.  If so then create a new structure
'                      rather than update a existing one
'
'*********************************************************************************

Dim bOnLoad As Byte
Dim bUpdated As Byte

Dim sAccount(10, 3) As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'*********************************************************************************

Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      SelectHelpTopic Me, "Financial Statement Structure"
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub


Private Sub cmdUpd_Click()
   Dim b As Byte
   Dim i As Integer
   bUpdated = True
   For i = 1 To 8
      If sAccount(i, 0) <> sAccount(i, 1) Then b = True
   Next
   If sAccount(i, 0) <> sAccount(i, 1) Then b = True
   If Not b Then
      MsgBox "The Accounts Haven't Changed.", _
         vbInformation, Caption
   Else
      UpdateAccounts
   End If
   
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillBoxes
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


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Dim b As Byte
   Dim bResponse As Byte
   Dim i As Integer
   Dim smsg As String
   For i = 1 To 8
      If sAccount(i, 0) <> sAccount(i, 1) Then b = True
   Next
   If sAccount(i, 0) <> sAccount(i, 1) Then b = True
   If b And Not bUpdated Then
      smsg = "The Structure Has Changed." & vbCrLf _
             & "Do You Want Exit Without Saving?"
      bResponse = MsgBox(smsg, ES_NOQUESTION, Caption)
      If bResponse = vbNo Then Cancel = True
   End If
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set diaGLe03a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Public Sub FillBoxes()
   Dim i As Integer
   Dim rdoGlm As rdoResultset
   On Error GoTo DiaErr1
   MouseCursor 13
   sSql = "SELECT * FROM GlmsTable WHERE COACCTREC=1"
   bSqlRows = GetDataSet(rdoGlm)
   If bSqlRows Then
      With rdoGlm
         txtAsst = "" & Trim(!COASSTACCT)
         lblAsst = "" & Trim(!COASSTDESC)
         sAccount(1, 0) = txtAsst
         
         txtLiab = "" & Trim(!COLIABACCT)
         lblLiab = "" & Trim(!COLIABDESC)
         sAccount(2, 0) = txtLiab
         
         txtEqty = "" & Trim(!COEQTYACCT)
         lblEqty = "" & Trim(!COEQTYDESC)
         sAccount(3, 0) = txtEqty
         
         txtIncm = "" & Trim(!COINCMACCT)
         lblIncm = "" & Trim(!COINCMDESC)
         sAccount(4, 0) = txtIncm
         
         txtCogs = "" & Trim(!COCOGSACCT)
         lblCogs = "" & Trim(!COCOGSDESC)
         sAccount(5, 0) = txtCogs
         
         txtExpn = "" & Trim(!COEXPNACCT)
         lblExpn = "" & Trim(!COEXPNDESC)
         sAccount(6, 0) = txtExpn
         
         txtOinc = "" & Trim(!COOINCACCT)
         lblOinc = "" & Trim(!COOINCDESC)
         sAccount(7, 0) = txtOinc
         
         txtOexp = "" & Trim(!COOEXPACCT)
         lblOexp = "" & Trim(!COOEXPDESC)
         sAccount(8, 0) = txtOexp
         
         txtFdtx = "" & Trim(!COFDTXACCT)
         lblFdtx = "" & Trim(!COFDTXDESC)
         sAccount(9, 0) = txtFdtx
         .Cancel
      End With
   End If
   For i = 1 To 8
      sAccount(i, 1) = sAccount(i, 0)
   Next
   sAccount(i, 1) = sAccount(i, 0)
   Set rdoGlm = Nothing
   MouseCursor 0
   Exit Sub
   
   DiaErr1:
   sProcName = "fillboxes"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub txtAsst_LostFocus()
   '1
   txtAsst = CheckLen(txtAsst, 12)
   If Len(txtAsst) = 0 Then
      Beep
      txtAsst = sAccount(1, 0)
   End If
   CheckName txtAsst, 1
   sAccount(1, 1) = txtAsst
   
End Sub


Private Sub txtCogs_LostFocus()
   '5
   txtCogs = CheckLen(txtCogs, 12)
   If Len(txtCogs) = 0 Then Beep
   If Len(txtCogs) Then CheckName txtCogs, 5
   sAccount(5, 1) = txtCogs
   
End Sub


Private Sub txtEqty_LostFocus()
   '3
   txtEqty = CheckLen(txtEqty, 12)
   If Len(txtEqty) = 0 Then
      Beep
      txtEqty = sAccount(3, 0)
   End If
   CheckName txtEqty, 3
   sAccount(3, 1) = txtEqty
   
End Sub


Private Sub txtExpn_LostFocus()
   '6
   txtExpn = CheckLen(txtExpn, 12)
   If Len(txtExpn) = 0 Then
      Beep
      txtExpn = sAccount(6, 0)
   End If
   CheckName txtExpn, 6
   sAccount(6, 1) = txtExpn
   
End Sub


Private Sub txtFdtx_LostFocus()
   '9
   txtFdtx = CheckLen(txtFdtx, 12)
   If Len(txtFdtx) = 0 Then Beep
   If Len(txtFdtx) Then CheckName txtFdtx, 9
   sAccount(9, 1) = txtFdtx
   
End Sub


Private Sub txtIncm_LostFocus()
   '4
   txtIncm = CheckLen(txtIncm, 12)
   If Len(txtIncm) = 0 Then
      Beep
      txtIncm = sAccount(4, 0)
   End If
   CheckName txtIncm, 4
   sAccount(4, 1) = txtIncm
   
End Sub


Private Sub txtLiab_LostFocus()
   '2
   txtLiab = CheckLen(txtLiab, 12)
   If Len(txtLiab) = 0 Then
      Beep
      txtLiab = sAccount(2, 0)
   End If
   CheckName txtLiab, 2
   sAccount(2, 1) = txtLiab
   
End Sub


Private Sub txtOexp_LostFocus()
   '8
   txtOexp = CheckLen(txtOexp, 12)
   If Len(txtOexp) = 0 Then Beep
   If Len(txtOexp) Then CheckName txtOexp, 8
   sAccount(8, 1) = txtOexp
   
End Sub


Private Sub txtOinc_LostFocus()
   '7
   txtOinc = CheckLen(txtOinc, 12)
   If Len(txtOinc) = 0 Then Beep
   If Len(txtOinc) Then CheckName txtOinc, 7
   sAccount(7, 1) = txtOinc
   
End Sub



Public Sub UpdateAccounts()
   Dim b As Byte
   Dim bResponse As Byte
   Dim i As Integer
   Dim smsg As String
   Dim rdoFirstRun As rdoResultset
   
   On Error GoTo DiaErr1
   smsg = "You Are About To Update Your Structure. Any " & vbCrLf _
          & "Blanked Accounts Will Mark Associate (Child)" & vbCrLf _
          & "Accounts As Inactive.  Do You Wish To " & vbCrLf _
          & "Continue With The Update?"
   bResponse = MsgBox(smsg, ES_NOQUESTION, Caption)
   If bResponse = vbYes Then
      MouseCursor 13
      cmdUpd.Enabled = False
      For i = 1 To 8
         sAccount(i, 2) = Compress(sAccount(i, 1))
      Next
      sAccount(i, 2) = Compress(sAccount(i, 1))
      
      '            For i = 1 To 9
      '                sSql = "SELECT GLACCTREF FROM GlacTable WHERE " _
      '                    & "GLACCTREF='" & sAccount(i, 2) & "' "
      '                RdoCon.Execute sSql, rdExecDirect
      '                If RdoCon.RowsAffected <> 0 Then
      '                    b = 1
      '                    Exit For
      '                End If
      '            Next
      '            If b = 1 Then
      '                MouseCursor 0
      '                MsgBox "The Number Selected For One Or More Accounts " & vbCrLf _
      '                    & "Is In Use By A Sub Account. Can't Update.", _
      '                    vbExclamation, Caption
      '                cmdUpd.Enabled = True
      '                Exit Sub
      '            End If
      MouseCursor 13
      On Error Resume Next
      
      ' Check if the GlmsTable is empty, If so then create a new record
      
      sSql = "SELECT COACCTREF FROM GlmsTable WHERE COACCTREC = 1"
      bSqlRows = GetDataSet(rdoFirstRun)
      
      
      If bSqlRows Then
         
         sSql = "UPDATE GlmsTable SET " _
                & "COASSTREF='" & sAccount(1, 2) & "'," _
                & "COASSTACCT='" & sAccount(1, 1) & "'," _
                & "COLIABREF='" & sAccount(2, 2) & "'," _
                & "COLIABACCT='" & sAccount(2, 1) & "'," _
                & "COEQTYREF='" & sAccount(3, 2) & "'," _
                & "COEQTYACCT='" & sAccount(3, 1) & "'," _
                & "COINCMREF='" & sAccount(4, 2) & "'," _
                & "COINCMACCT='" & sAccount(4, 1) & "'," _
                & "COCOGSREF='" & sAccount(5, 2) & "'," _
                & "COCOGSACCT='" & sAccount(5, 1) & "'," _
                & "COEXPNREF='" & sAccount(6, 2) & "'," _
                & "COEXPNACCT='" & sAccount(6, 1) & "'," _
                & "COOINCREF='" & sAccount(7, 2) & "'," _
                & "COOINCACCT='" & sAccount(7, 1) & "'," _
                & "COOEXPREF='" & sAccount(8, 2) & "'," _
                & "COOEXPACCT='" & sAccount(8, 1) & "'," _
                & "COFDTXREF='" & sAccount(9, 2) & "'," _
                & "COFDTXACCT='" & sAccount(9, 1) & "' " _
                & "WHERE COACCTREC=1 "
         
         
      Else
         
         
         
         sSql = "INSERT INTO GlmsTable (COASSTREF,COASSTACCT,COLIABREF,COLIABACCT," _
                & "COEQTYREF,COEQTYACCT,COINCMREF,COINCMACCT,COCOGSREF,COCOGSACCT," _
                & "COEXPNREF,COEXPNACCT,COOINCREF,COOINCACCT,COOEXPREF,COOEXPACCT," _
                & "COFDTXREF,COFDTXACCT,COACCTREC) VALUES('" _
                & sAccount(1, 2) & "','" & sAccount(1, 1) & "','" _
                & sAccount(2, 2) & "','" & sAccount(2, 1) & "','" _
                & sAccount(3, 2) & "','" & sAccount(3, 1) & "','" _
                & sAccount(4, 2) & "','" & sAccount(4, 1) & "','" _
                & sAccount(5, 2) & "','" & sAccount(5, 1) & "','" _
                & sAccount(6, 2) & "','" & sAccount(6, 2) & "','" _
                & sAccount(7, 2) & "','" & sAccount(7, 2) & "','" _
                & sAccount(8, 2) & "','" & sAccount(8, 2) & "','" _
                & sAccount(9, 2) & "','" & sAccount(9, 2) & "',1)"
         
      End If
      Set rdoFirstRun = Nothing
      
      On Error GoTo 0
      RdoCon.Execute sSql, rdExecDirect
      If RdoCon.RowsAffected > 0 Then
         For i = 1 To 9
            If sAccount(i, 0) <> sAccount(i, 1) Then
               If sAccount(i, 1) = "" Then b = 1 Else b = 0
               sSql = "UPDATE GlacTable SET " _
                      & "GLMASTER='" & sAccount(i, 2) & "'," _
                      & "GLINACTIVE=" & b & " " _
                      & "WHERE GLMASTER='" & sAccount(i, 0) & "' "
               RdoCon.Execute sSql, rdExecDirect
            End If
         Next
         MouseCursor 0
         bUpdated = True
         MsgBox "Structure Successfully Updated.", _
            vbInformation, Caption
         Unload Me
      Else
         MouseCursor 0
         MsgBox "Couldn't Update Structure.", _
            vbExclamation, Caption
         bUpdated = False
         cmdUpd.Enabled = True
      End If
   Else
      CancelTrans
      cmdUpd.Enabled = True
   End If
   Exit Sub
   
   DiaErr1:
   sProcName = "updateacc"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Public Sub CheckName(sAccounts As String, Index As Integer)
   Dim b As Byte
   Dim i As Integer
   On Error Resume Next
   For i = 0 To 8
      If i <> Index Then
         If sAccount(i, 1) = sAccounts Then b = 1
      End If
   Next
   If b = 1 Then
      Beep
      Select Case Index
         Case 1
            txtAsst = sAccount(1, 1)
         Case 2
            txtLiab = sAccount(2, 1)
         Case 3
            txtEqty = sAccount(3, 1)
         Case 4
            txtIncm = sAccount(4, 1)
         Case 5
            txtCogs = sAccount(5, 1)
         Case 6
            txtExpn = sAccount(6, 1)
         Case 7
            txtOinc = sAccount(7, 1)
         Case 8
            txtOexp = sAccount(8, 1)
         Case Else
            txtFdtx = sAccount(8, 1)
      End Select
   End If
   
End Sub
