VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form diaGlchg
   BorderStyle = 3 'Fixed Dialog
   Caption = "Change An Account Number"
   ClientHeight = 2220
   ClientLeft = 45
   ClientTop = 330
   ClientWidth = 5175
   ClipControls = 0 'False
   ControlBox = 0 'False
   LinkTopic = "Form1"
   MaxButton = 0 'False
   MDIChild = -1 'True
   MinButton = 0 'False
   ScaleHeight = 2220
   ScaleWidth = 5175
   ShowInTaskbar = 0 'False
   Begin ComctlLib.ProgressBar Prg1
      Height = 255
      Left = 2040
      TabIndex = 8
      Top = 1320
      Visible = 0 'False
      Width = 2055
      _ExtentX = 3625
      _ExtentY = 450
      _Version = 327682
      Appearance = 1
   End
   Begin VB.CommandButton cmdUpd
      Caption = "&Update"
      Height = 315
      Left = 4200
      TabIndex = 2
      ToolTipText = "Change Account Number And Update References"
      Top = 600
      Width = 875
   End
   Begin VB.TextBox txtAct
      Height = 285
      Left = 2040
      TabIndex = 0
      Tag = "3"
      Top = 960
      Width = 1815
   End
   Begin VB.CommandButton cmdCan
      Cancel = -1 'True
      Caption = "Close"
      Height = 435
      Left = 4200
      TabIndex = 4
      TabStop = 0 'False
      Top = 120
      Width = 875
   End
   Begin Threed.SSRibbon cmdHlp
      Height = 225
      Left = 0
      TabIndex = 5
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
      PictureUp = "diaGlchg.frx":0000
      PictureDn = "diaGlchg.frx":0146
   End
   Begin ResizeLibCtl.ReSize ReSize1
      Left = 3240
      Top = 0
      _Version = 196615
      _ExtentX = 741
      _ExtentY = 741
      _StockProps = 0
      Enabled = -1 'True
      FormMinWidth = 0
      FormMinHeight = 0
      FormDesignHeight = 2220
      FormDesignWidth = 5175
   End
   Begin VB.Label lblUpd
      Caption = "Updating"
      BeginProperty Font
      Name = "MS Sans Serif"
      Size = 8.25
      Charset = 0
      Weight = 400
      Underline = 0 'False
      Italic = -1 'True
      Strikethrough = 0 'False
      EndProperty
      ForeColor = &H00800000&
      Height = 255
      Left = 120
      TabIndex = 7
      Top = 1320
      Visible = 0 'False
      Width = 1935
   End
   Begin VB.Label lblAct
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      Height = 285
      Left = 2040
      TabIndex = 6
      Top = 600
      Width = 1815
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "To Account Number"
      Height = 375
      Index = 1
      Left = 120
      TabIndex = 3
      Top = 960
      Width = 1815
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Change From"
      Height = 255
      Index = 0
      Left = 120
      TabIndex = 1
      Top = 600
      Width = 1935
   End
End
Attribute VB_Name = "diaGlchg"
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

Dim bOnLoad As Byte
Dim bGoodAcct As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      SelectHelpTopic Me, "No Subject Help"
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub


Private Sub cmdUpd_Click()
   If Len(Trim(txtAct)) And txtAct <> lblAct Then
      bGoodAcct = GetAccount()
      If bGoodAcct = 1 Then
         UpdateAccount
      Else
         MsgBox "That Account Is Already Included.", _
            vbExclamation, Caption
      End If
   End If
   
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      bOnLoad = False
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   SetDiaPos Me, ES_DONTLIST
   Move diaGLe01a.Left + 400, diaGLe01a.Top + 1400
   FormatControls
   sCurrForm = Caption
   lblAct = diaGLe01a.cmbAct
   bOnLoad = True
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   diaGLe01a.optChg.Value = vbUnchecked
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Set diaGlchg = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Public Sub FillCombo()
   On Error GoTo DiaErr1
   Exit Sub
   
   DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   
End Sub

Private Sub txtAct_LostFocus()
   txtAct = CheckLen(txtAct, 12)
   
End Sub



Public Function GetAccount() As Byte
   Dim rdoAct As rdoResultset
   Dim sAccount As String
   
   On Error GoTo DiaErr1
   sAccount = Compress(txtAct)
   sSql = "SELECT GLACCTREF,GLACCTNO FROM GlacTable " _
          & "WHERE GLACCTREF='" & sAccount & "' "
   bSqlRows = GetDataSet(rdoAct, ES_FORWARD)
   If bSqlRows Then
      GetAccount = 0
   Else
      GetAccount = 1
   End If
   Set rdoAct = Nothing
   Exit Function
   
   DiaErr1:
   sProcName = "getaccoun"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Public Sub UpdateAccount()
   Dim i As Integer
   Dim bResponse As Byte
   Dim smsg As String
   Dim sNewAccount As String
   Dim sOldAccount As String
   
   'On Error GoTo DiaErr1
   sNewAccount = Compress(txtAct)
   sOldAccount = Compress(lblAct)
   smsg = "Do You Really Want To Change " & Trim(lblAct) & vbCrLf _
          & "And Update To Account Number " & txtAct & "... "
   bResponse = MsgBox(smsg, ES_NOQUESTION, Caption)
   If bResponse = vbYes Then
      cmdUpd.Enabled = False
      MouseCursor 13
      Err = 0
      On Error Resume Next
      RdoCon.BeginTrans
      
      lblUpd.Visible = True
      Prg1.Visible = True
      lblUpd = "Updating Base"
      Prg1.Value = 10
      lblUpd.Refresh
      'Base Table
      sSql = "UPDATE GlacTable SET GLACCTREF='" & sNewAccount & "'," _
             & "GLACCTNO='" & txtAct & "' " _
             & "WHERE GLACCTREF='" & sOldAccount & "'"
      RdoCon.Execute sSql, rdExecDirect
      
      sSql = "UPDATE GlacTable SET GLMASTER='" & sNewAccount & "' " _
             & "WHERE GLMASTER='" & sOldAccount & "'"
      RdoCon.Execute sSql, rdExecDirect
      
      'Product Codes
      lblUpd = "Updating Product Codes"
      Prg1.Value = 20
      lblUpd.Refresh
      sSql = "UPDATE PcodTable SET PCREVACCT='" & sNewAccount & "' " _
             & "WHERE PCREVACCT='" & sOldAccount & "'"
      RdoCon.Execute sSql, rdExecDirect
      
      sSql = "UPDATE PcodTable SET PCDISCACCT='" & sNewAccount & "' " _
             & "WHERE PCDISCACCT='" & sOldAccount & "'"
      RdoCon.Execute sSql, rdExecDirect
      
      sSql = "UPDATE PcodTable SET PCDREVXFERAC='" & sNewAccount & "' " _
             & "WHERE PCDREVXFERAC='" & sOldAccount & "'"
      RdoCon.Execute sSql, rdExecDirect
      
      Prg1.Value = 30
      sSql = "UPDATE PcodTable SET PCDCGSXFERAC='" & sNewAccount & "' " _
             & "WHERE PCDCGSXFERAC='" & sOldAccount & "'"
      RdoCon.Execute sSql, rdExecDirect
      
      sSql = "UPDATE PcodTable SET PCINVEXAC='" & sNewAccount & "' " _
             & "WHERE PCINVEXAC='" & sOldAccount & "'"
      RdoCon.Execute sSql, rdExecDirect
      
      Prg1.Value = 40
      sSql = "UPDATE PcodTable SET PCCGSAC='" & sNewAccount & "' " _
             & "WHERE PCCGSAC='" & sOldAccount & "'"
      RdoCon.Execute sSql, rdExecDirect
      
      'Parts
      lblUpd = "Updating Parts"
      Prg1.Value = 50
      lblUpd.Refresh
      sSql = "UPDATE PartTable SET PAACCTNO='" & sNewAccount & "' " _
             & "WHERE PAACCTNO='" & sOldAccount & "'"
      RdoCon.Execute sSql, rdExecDirect
      
      sSql = "UPDATE PartTable SET PAREVACCT='" & sNewAccount & "' " _
             & "WHERE PAREVACCT='" & sOldAccount & "'"
      RdoCon.Execute sSql, rdExecDirect
      
      Prg1.Value = 60
      sSql = "UPDATE PartTable SET PACGSACCT='" & sNewAccount & "' " _
             & "WHERE PACGSACCT='" & sOldAccount & "'"
      RdoCon.Execute sSql, rdExecDirect
      
      sSql = "UPDATE PartTable SET PADISACCT='" & sNewAccount & "' " _
             & "WHERE PADISACCT='" & sOldAccount & "'"
      RdoCon.Execute sSql, rdExecDirect
      
      Prg1.Value = 80
      sSql = "UPDATE PartTable SET PATFRREVACCT='" & sNewAccount & "' " _
             & "WHERE PATFRREVACCT='" & sOldAccount & "'"
      RdoCon.Execute sSql, rdExecDirect
      
      sSql = "UPDATE PartTable SET PATFRCGSACCT='" & sNewAccount & "' " _
             & "WHERE PATFRCGSACCT='" & sOldAccount & "'"
      RdoCon.Execute sSql, rdExecDirect
      
      sSql = "UPDATE PartTable SET PAREJACCT='" & sNewAccount & "' " _
             & "WHERE PAREJACCT='" & sOldAccount & "'"
      RdoCon.Execute sSql, rdExecDirect
      
      Prg1.Value = 100
      MouseCursor 0
      If Err <> 0 Then
         RdoCon.CommitTrans
         MsgBox "Account Number Was Successfully Changed.", _
            vbInformation, Caption
         For i = 0 To diaGLe01a.cmbAct.ListCount - 1
            If diaGLe01a.cmbAct.List(i) = lblAct Then
               diaGLe01a.cmbAct.List(i) = txtAct
               Exit For
            End If
         Next
         diaGLe01a.cmbAct = txtAct
         Prg1.Visible = False
         lblUpd.Visible = False
         Unload Me
      Else
         RdoCon.RollbackTrans
         RdoCon.CommitTrans
         MsgBox "Could Not Change The Existing Account.", _
            vbExclamation, Caption
         diaGLe01a.cmbAct = txtAct
         Prg1.Visible = False
         lblUpd.Visible = False
         Unload Me
      End If
   Else
      CancelTrans
   End If
   Exit Sub
   
   DiaErr1:
   sProcName = "updateacc"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   
End Sub
