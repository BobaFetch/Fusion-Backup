VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaAPp22a
   BorderStyle = 1 'Fixed Single
   Caption = "Reconciliation Report"
   ClientHeight = 3300
   ClientLeft = 45
   ClientTop = 330
   ClientWidth = 6540
   ControlBox = 0 'False
   LinkTopic = "Form1"
   MaxButton = 0 'False
   MDIChild = -1 'True
   MinButton = 0 'False
   ScaleHeight = 3300
   ScaleWidth = 6540
   Begin VB.ComboBox cmbAct
      Height = 315
      Left = 960
      TabIndex = 0
      Top = 840
      Width = 1575
   End
   Begin VB.CheckBox optSum
      Caption = "___"
      ForeColor = &H8000000F&
      Height = 195
      Left = 2040
      TabIndex = 3
      Top = 2880
      Width = 735
   End
   Begin VB.ComboBox txteDte
      Height = 315
      Left = 2040
      TabIndex = 2
      Tag = "4"
      Top = 2400
      Width = 1095
   End
   Begin VB.Frame fraPrn
      BorderStyle = 0 'None
      Height = 495
      Left = 5400
      TabIndex = 7
      Top = 360
      Width = 1095
      Begin VB.CommandButton optPrn
         Height = 330
         Left = 560
         Picture = "diaGlp15a.frx":0000
         Style = 1 'Graphical
         TabIndex = 5
         ToolTipText = "Print The Report"
         Top = 120
         UseMaskColor = -1 'True
         Width = 495
      End
      Begin VB.CommandButton optDis
         Height = 330
         Left = 0
         Picture = "diaGlp15a.frx":018A
         Style = 1 'Graphical
         TabIndex = 4
         ToolTipText = "Display The Report"
         Top = 120
         UseMaskColor = -1 'True
         Width = 495
      End
   End
   Begin VB.CommandButton cmdCan
      Cancel = -1 'True
      Caption = "Close"
      Height = 360
      Left = 5400
      TabIndex = 6
      TabStop = 0 'False
      Top = 0
      Width = 1065
   End
   Begin VB.ComboBox txtsDte
      Height = 315
      Left = 2040
      TabIndex = 1
      Tag = "4"
      Top = 2040
      Width = 1095
   End
   Begin Threed.SSRibbon cmdHlp
      Height = 225
      Left = 0
      TabIndex = 8
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
      PictureUp = "diaGlp15a.frx":0308
      PictureDn = "diaGlp15a.frx":044E
   End
   Begin ResizeLibCtl.ReSize ReSize1
      Left = 5760
      Top = 1800
      _Version = 196615
      _ExtentX = 741
      _ExtentY = 741
      _StockProps = 0
      Enabled = -1 'True
      FormMinWidth = 0
      FormMinHeight = 0
      FormDesignHeight = 3300
      FormDesignWidth = 6540
   End
   Begin Threed.SSRibbon ShowPrinters
      Height = 255
      Left = 360
      TabIndex = 9
      ToolTipText = "Show System Printers"
      Top = 0
      Width = 255
      _Version = 65536
      _ExtentX = 450
      _ExtentY = 450
      _StockProps = 65
      BackColor = 12632256
      GroupAllowAllUp = -1 'True
      RoundedCorners = 0 'False
      BevelWidth = 0
      Outline = 0 'False
      PictureUp = "diaGlp15a.frx":0594
      PictureDn = "diaGlp15a.frx":06DA
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Account"
      Height = 255
      Index = 0
      Left = 240
      TabIndex = 16
      Top = 840
      Width = 1455
   End
   Begin VB.Label lblDsc
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      Height = 285
      Index = 0
      Left = 960
      TabIndex = 15
      Top = 1200
      Width = 2775
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Summary Only?"
      Height = 285
      Index = 11
      Left = 360
      TabIndex = 14
      Top = 2880
      Width = 1575
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Ending"
      Height = 285
      Index = 10
      Left = 600
      TabIndex = 13
      Top = 2400
      Width = 1545
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Beginning"
      Height = 285
      Index = 9
      Left = 600
      TabIndex = 12
      Top = 2100
      Width = 1545
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Reconciliation Period"
      Height = 285
      Index = 4
      Left = 240
      TabIndex = 11
      Top = 1800
      Width = 2385
   End
   Begin VB.Label lblPrinter
      Appearance = 0 'Flat
      BorderStyle = 1 'Fixed Single
      Caption = "Default Printer"
      ForeColor = &H00800000&
      Height = 255
      Left = 720
      TabIndex = 10
      Top = 0
      Width = 2760
   End
End
Attribute VB_Name = "diaAPp22a"
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
' diaARp06a- Account Reconciliation Report
'
' Created: 12/10/03 (JcW!)
'
'
'*********************************************************************************
Dim bOnLoad As Byte
Dim bGoodVendor As Boolean

' Accounts
Dim sCrCashAcct As String
Dim sCrDiscAcct As String
Dim sCrExpAcct As String
Dim sSJARAcct As String
Dim sCrRevAcct As String
Dim sCrCommAcct As String

'keys
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
      SelectHelpTopic Me, "Reports"
      cmdHlp = False
      MouseCursor 0
   End If
End Sub

Private Sub cmbAct_Click()
   lblDsc(0) = UpdateActDesc(cmbAct)
End Sub

Private Sub cmbAct_LostFocus()
   cmbAct = CheckLen(cmbAct, 12)
   If Trim(cmbAct) <> "" Then
      lblDsc(0) = UpdateActDesc(cmbAct)
   Else
      lblDsc(0) = "*** Invalid Account Number ***"
   End If
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillAccounts
      GetOptions
      bOnLoad = False
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   SetDiaPos Me
   FormatControls
   bOnLoad = True
   txtsDte = Format(Now, "mm/01/yy")
   txteDte = Format(Now, "mm/dd/yy")
End Sub

Public Sub FillAccounts()
   ' Fill account combo
   ' Need to add account descriptions
   Dim RdoAct As rdoResultset
   Dim b As Byte
   
   On Error GoTo DiaErr1
   
   b = GetCashAccounts()
   If b = 3 Then
      MouseCursor 0
      MsgBox "One Or More Cash Accounts Are Not Active." & vbCr _
         & "Please Set All Cash Accounts In The " & vbCr _
         & "System Setup, Administration Section.", _
         vbExclamation, Caption
      Sleep 500
      Unload Me
      Exit Sub
   End If
   
   ' Accounts
   sSql = "SELECT GLACCTNO FROM GlacTable WHERE GLCASH=1"
   bSqlRows = GetDataSet(RdoAct)
   If bSqlRows Then
      With RdoAct
         While Not .EOF
            AddComboStr cmbAct.hwnd, "" & Trim(!GLACCTNO)
            .MoveNext
         Wend
      End With
      cmbAct.ListIndex = 0
      lblDsc(0) = UpdateActDesc(cmbAct)
   Else
      ' Multiple cash accounts not found so use the default cash account
      cmbAct.Enabled = False
      cmbAct = sCrCashAcct
      lblDsc(0) = UpdateActDesc(cmbAct)
   End If
   
   
   Set RdoAct = Nothing
   
   ' Other Accounts
   'sSql = "Qry_FillLowAccounts"
   'bSqlRows = GetDataSet(RdoAct, ES_FORWARD)
   'If bSqlRows Then
   '    With RdoAct
   '        Do Until .EOF
   '            AddComboStr cmbSerAct.hwnd, "" & Trim(!GLACCTNO)
   '            AddComboStr cmbIntAct.hwnd, "" & Trim(!GLACCTNO)
   '            .MoveNext
   '        Loop
   '    End With
   '    cmbSerAct.ListIndex = 0
   '    cmbIntAct.ListIndex = 0
   'End If
   
   Set RdoAct = Nothing
   Exit Sub
   
   DiaErr1:
   sProcName = "FillAcounts"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

'Local Errors - from FillCombo
'Flash with 1 = no accounts nec, 2 = OK, 3 = Not enough accounts

Public Function GetCashAccounts() As Byte
   Dim rdocsh As rdoResultset
   Dim i As Integer
   Dim b As Byte
   On Error GoTo DiaErr1
   sSql = "SELECT COGLVERIFY,COCRCASHACCT,COCRDISCACCT,COSJARACCT," _
          & "COCRCOMMACCT,COCRREVACCT,COCREXPACCT FROM ComnTable WHERE COREF=1"
   bSqlRows = GetDataSet(rdocsh, ES_FORWARD)
   sProcName = "getcashacct"
   If bSqlRows Then
      With rdocsh
         For i = 1 To 6
            If "" & Trim(.rdoColumns(i)) = "" Then
               b = 1
               Exit For
            End If
         Next
         sCrCashAcct = "" & Trim(!COCRCASHACCT)
         sCrDiscAcct = "" & Trim(!COCRDISCACCT)
         sSJARAcct = "" & Trim(!COSJARACCT)
         sCrCommAcct = "" & Trim(!COCRCOMMACCT)
         sCrRevAcct = "" & Trim(!COCRREVACCT)
         sCrExpAcct = "" & Trim(!COCREXPACCT)
         .Cancel
         If b = 1 Then GetCashAccounts = 3 Else GetCashAccounts = 2
      End With
   Else
      GetCashAccounts = 0
   End If
   Set rdocsh = Nothing
   
   Exit Function
   DiaErr1:
   sProcName = "GetCashAccounts"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Function

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   FormUnload
   SaveOptions
   Set diaAPp22a = Nothing
End Sub

Public Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Private Sub PrintReport()
   Dim sCustomReport As String
   MouseCursor 13
   On Error GoTo DiaErr1
   
   If Left(lblDsc(0), 5) = "*** I" Then
      MsgBox "Enter A Valid Account.", vbExclamation, Caption
      MouseCursor 0
      Exit Sub
   End If
   
   SetMdiReportsize MdiSect
   
   sCustomReport = GetCustomReport("finap22a.rpt")
   MdiSect.crw.ReportFileName = sReportPath & sCustomReport
   
   If Trim(lblDsc(0)) <> "*** Invalid Account Number ***" Then
      sSql = "{GlacTable.GLACCTREF} = '" & Compress(cmbAct) & "'"
   Else
      Exit Sub
      MsgBox "Enter a Valid Account", vbExclamation, Caption
   End If
   
   MdiSect.crw.Formulas(0) = "CompanyName='" & sFacility & "'"
   MdiSect.crw.Formulas(1) = "RequestedBy='Requested By ESI'"
   MdiSect.crw.Formulas(2) = "Title1='Reconciliation Report'"
   MdiSect.crw.Formulas(3) = "Title2='Period Beginning " & Trim(txtsDte) & "   Ending " & Trim(txteDte) & "'"
   
   If optSum.Value = vbChecked Then
      MdiSect.crw.Formulas(4) = "subreport='1'"
   Else
      MdiSect.crw.Formulas(4) = "subreport='0'"
   End If
   
   If Trim(txtsDte) <> "" Then
      MdiSect.crw.Formulas(5) = "Beginning=cdate('" & txtsDte & "')"
   End If
   
   If Trim(txteDte) <> "" Then
      MdiSect.crw.Formulas(6) = "ending=cdate('" & txteDte & "')"
   End If
   
   MdiSect.crw.Formulas(7) = "Account='Account: " & cmbAct & "'"
   MdiSect.crw.SelectionFormula = sSql
   
   SetCrystalAction Me
   MouseCursor 0
   Exit Sub
   
   DiaErr1:
   sProcName = "PrintReport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub optDis_Click()
   PrintReport
End Sub

Private Sub optPrn_Click()
   PrintReport
End Sub

Private Sub txtEDte_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtEdte_LostFocus()
   txteDte = CheckDate(txteDte)
End Sub

Private Sub txtSDte_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtSDte_LostFocus()
   txtsDte = CheckDate(txtsDte)
End Sub


Public Sub SaveOptions()
   Dim sOptions As String
   'Save by Menu Option
   sOptions = RTrim(optSum.Value)
   SaveSetting "Esi2000", "EsiFina", Me.Name, Trim(sOptions)
   SaveSetting "Esi2000", "EsiFina", Me.Name & "_Printer", lblPrinter
End Sub

Public Sub GetOptions()
   Dim sOptions As String
   sOptions = GetSetting("Esi2000", "EsiFina", Me.Name, sOptions)
   If Len(Trim(sOptions)) > 0 Then
      optSum.Value = Val(Left(sOptions, 1))
   Else
      optSum.Value = vbUnchecked
   End If
   lblPrinter = GetSetting("Esi2000", "EsiFina", Me.Name & "_Printer", lblPrinter)
   If lblPrinter = "" Then
      lblPrinter = "Default Printer"
   End If
End Sub
