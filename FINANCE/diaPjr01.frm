VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaJRp01a
   BorderStyle = 3 'Fixed Dialog
   Caption = "Journals (Report)"
   ClientHeight = 3825
   ClientLeft = 2115
   ClientTop = 1125
   ClientWidth = 7290
   ClipControls = 0 'False
   ControlBox = 0 'False
   LinkTopic = "Form1"
   MDIChild = -1 'True
   MinButton = 0 'False
   PaletteMode = 1 'UseZOrder
   ScaleHeight = 3825
   ScaleWidth = 7290
   ShowInTaskbar = 0 'False
   Tag = "4"
   Begin VB.ComboBox cmbTyp
      ForeColor = &H00800000&
      Height = 315
      Left = 1320
      TabIndex = 1
      Tag = "8"
      ToolTipText = "Select GL Type From List"
      Top = 960
      Width = 1095
   End
   Begin VB.ComboBox cmbFyr
      ForeColor = &H00800000&
      Height = 315
      Left = 1320
      Sorted = -1 'True
      TabIndex = 0
      Tag = "8"
      ToolTipText = "Select Fiscal Year"
      Top = 600
      Width = 1095
   End
   Begin VB.ComboBox cmbJrn
      Height = 315
      Left = 1320
      Sorted = -1 'True
      TabIndex = 2
      Tag = "3"
      ToolTipText = "Select A Journal From The List"
      Top = 1560
      Width = 1775
   End
   Begin VB.CommandButton cmdCan
      Cancel = -1 'True
      Caption = "Close"
      Height = 360
      Left = 6120
      TabIndex = 7
      TabStop = 0 'False
      Top = 0
      Width = 1065
   End
   Begin VB.Frame fraPrn
      BorderStyle = 0 'None
      Height = 495
      Left = 6120
      TabIndex = 4
      Top = 360
      Width = 1095
      Begin VB.CommandButton optDis
         Height = 330
         Left = 0
         Style = 1 'Graphical
         TabIndex = 5
         ToolTipText = "Display The Report"
         Top = 120
         UseMaskColor = -1 'True
         Width = 495
      End
      Begin VB.CommandButton optPrn
         Height = 330
         Left = 560
         Style = 1 'Graphical
         TabIndex = 6
         ToolTipText = "Print The Report"
         Top = 120
         UseMaskColor = -1 'True
         Width = 495
      End
   End
   Begin Threed.SSRibbon cmdHlp
      Height = 225
      Left = 0
      TabIndex = 3
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
      PictureUp = "diaPjr01.frx":0000
      PictureDn = "diaPjr01.frx":0146
   End
   Begin ResizeLibCtl.ReSize ReSize1
      Left = 3600
      Top = 0
      _Version = 196615
      _ExtentX = 741
      _ExtentY = 741
      _StockProps = 0
      Enabled = -1 'True
      FormMinWidth = 0
      FormMinHeight = 0
      FormDesignHeight = 3825
      FormDesignWidth = 7290
   End
   Begin Threed.SSRibbon ShowPrinters
      Height = 255
      Left = 360
      TabIndex = 18
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
      PictureUp = "diaPjr01.frx":028C
      PictureDn = "diaPjr01.frx":03D2
   End
   Begin VB.Label lblEnd
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      Height = 285
      Left = 4440
      TabIndex = 23
      Top = 2280
      Width = 855
   End
   Begin VB.Label lblStart
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      Height = 285
      Left = 1320
      TabIndex = 22
      Top = 2280
      Width = 855
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "End"
      Height = 255
      Index = 4
      Left = 3600
      TabIndex = 21
      Top = 2280
      Width = 735
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Start"
      Height = 255
      Index = 3
      Left = 240
      TabIndex = 20
      Top = 2280
      Width = 975
   End
   Begin VB.Label lblPrinter
      Appearance = 0 'Flat
      BorderStyle = 1 'Fixed Single
      ForeColor = &H00800000&
      Height = 255
      Left = 720
      TabIndex = 19
      Top = 0
      Width = 2760
   End
   Begin VB.Label lblkind
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      Height = 285
      Left = 2520
      TabIndex = 17
      Top = 960
      Width = 2775
   End
   Begin VB.Label P
      BackStyle = 0 'Transparent
      Caption = "Type"
      Height = 285
      Index = 2
      Left = 240
      TabIndex = 16
      Top = 960
      Width = 825
   End
   Begin VB.Label lblNum
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      Height = 285
      Left = 5400
      TabIndex = 15
      Top = 2640
      Visible = 0 'False
      Width = 495
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Closed"
      Height = 255
      Index = 1
      Left = 3600
      TabIndex = 14
      Top = 2640
      Width = 855
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Opened"
      Height = 255
      Index = 0
      Left = 240
      TabIndex = 13
      Top = 2640
      Width = 855
   End
   Begin VB.Label lblOpen
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      Height = 285
      Left = 1320
      TabIndex = 12
      Top = 2640
      Width = 855
   End
   Begin VB.Label lblClose
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      Height = 285
      Left = 4440
      TabIndex = 11
      Top = 2640
      Width = 855
   End
   Begin VB.Label P
      BackStyle = 0 'Transparent
      Caption = "Fiscal Year"
      Height = 285
      Index = 1
      Left = 240
      TabIndex = 10
      Top = 600
      Width = 1425
   End
   Begin VB.Label lblDsc
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      Height = 285
      Left = 1320
      TabIndex = 9
      Top = 1920
      Width = 2775
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Journal "
      Height = 255
      Index = 2
      Left = 240
      TabIndex = 8
      Top = 1560
      Width = 1335
   End
End
Attribute VB_Name = "diaJRp01a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of                     ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***

'See the UpdateTables prodecure for database revisions

'**************************************************************************************
' Form: diaPjr01a - Display Journal Reports
'
' Notes: This form takes the places of all MCS journal viewing/reporting programs
'
' Created:  (cjs)
' Modified:
'   06/05/01 (nth) Redesigned window layout and included the SJ.
'   06/17/01 (nth) Added to INVCANCELED to sales journal selection formula.
'   11/11/02 (nth) Updated XC and PJ Journals.
'   11/14/02 (nth) Updated CR Journal.
'   12/11/02 (nth) Updated CC Journal.
'   04/04/03 (nth) Correctly display voided checks.
'   06/05/03 (nth) Removed IJ,TJ,and PR journal types.
'   09/18/03 (nth) Allow credit and debit memos to correctly display on SJ.
'   04/01/04 (nth) Canceled invoices on sales journal.
'   04/06/04 (nth) Add DCHEAD formula to CR journal.
'   08/16/04 (nth) Added printer saveoptions and getoptions
'
'**************************************************************************************

Option Explicit

Dim bOnLoad As Byte
Dim bCancel As Byte
Dim bGoodId As Byte
Dim iFyear As Integer
Dim iJrnNo As Integer

Dim sKind(12) As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'**************************************************************************************

Private Sub cmbFyr_Click()
   FillJournals
End Sub

Private Sub cmbFyr_LostFocus()
   Dim I As Integer
   Dim b As Byte
   
   If Not bCancel = 0 Then
      cmbFyr = CheckLen(cmbFyr, 4)
      cmbFyr = Format(Abs(Val(cmbFyr)), "0000")
      For I = 0 To cmbFyr.ListCount - 1
         If cmbFyr = cmbFyr.List(I) Then b = 1
      Next
      If b = 0 Then
         Beep
         cmbFyr = Format(Now, "yyyy")
      End If
      FillJournals
   End If
End Sub


Private Sub cmbjrn_Click()
   bGoodId = GetJrnId()
End Sub

Private Sub cmbjrn_LostFocus()
   cmbJrn = CheckLen(cmbJrn, 12)
   If bCancel = 0 Then bGoodId = GetJrnId()
End Sub

Private Sub cmbTyp_Click()
   FillJournals
   If cmbTyp.ListIndex > 0 Then lblkind = sKind(cmbTyp.ListIndex)
End Sub

Private Sub cmbTyp_LostFocus()
   If bCancel = 0 Then
      cmbTyp = CheckLen(cmbTyp, 3)
      If Trim(cmbTyp) = "" Then cmbTyp = "ALL"
      FillJournals
   End If
End Sub

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer _
                             , X As Single, Y As Single)
   bCancel = 1
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      SelectHelpTopic Me, "Reports"
      MouseCursor 0
      cmdHlp = False
   End If
End Sub

Private Sub FillCombo()
   Dim RdoPst As rdoResultset
   Dim I As Integer
   
   On Error GoTo DiaErr1
   
   sSql = "SELECT FYYEAR FROM GlfyTable "
   bSqlRows = GetDataSet(RdoPst, ES_FORWARD)
   
   If bSqlRows Then
      With RdoPst
         Do Until .EOF
            If Not IsNull(.rdoColumns(0)) Then _
                          AddComboStr cmbFyr.hWnd, Format(.rdoColumns(0), "0000")
            .MoveNext
         Loop
         .Cancel
      End With
   End If
   
   AddComboStr cmbTyp.hWnd, "ALL"
   sKind(0) = "ALL"
   AddComboStr cmbTyp.hWnd, "SJ"
   sKind(1) = "Sales"
   AddComboStr cmbTyp.hWnd, "PJ"
   sKind(2) = "Purchases"
   AddComboStr cmbTyp.hWnd, "CR"
   sKind(3) = "Cash Receipts"
   AddComboStr cmbTyp.hWnd, "CC"
   sKind(4) = "Disp-Computer Checks"
   AddComboStr cmbTyp.hWnd, "XC"
   sKind(5) = "Disp-External Checks"
   AddComboStr cmbTyp.hWnd, "TJ"
   sKind(6) = "Time Charges"
   
   'AddComboStr cmbTyp.hWnd, "PL"
   'sKind(6) = "Payroll Labor"
   'addComboStr cmbTyp.hWnd, "PD"
   'sKind(7) = "Disp-Payroll"
   'AddComboStr cmbTyp.hWnd, "TJ"
   'sKind(8) = "Time Journal"
   
   cmbTyp = "ALL"
   lblkind = sKind(0)
   If cmbFyr.ListCount > 0 Then
      cmbFyr = Format(Now, "yyyy")
      FillJournals
   End If
   Exit Sub
   
   DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub Form_Activate()
   On Error Resume Next
   MdiSect.lblBotPanel = Caption
   If bOnLoad = 1 Then
      MouseCursor 13
      FillCombo
      bOnLoad = 0
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   SetDiaPos Me
   FormatControls
   sCurrForm = Caption
   lblDsc.ForeColor = Me.ForeColor
   GetOptions
   optPrn.Picture = Resources.imgPrn.Picture
   optDis.Picture = Resources.imgDis.Picture
   bOnLoad = 1
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveOptions
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set diaJRp01a = Nothing
End Sub

Private Sub GetLabor(sAcct As String, sDesc As String)
   Dim RdoLab As rdoResultset
   On Error GoTo DiaErr1
   sSql = " SELECT GLACCTNO,GLDESCR FROM ComnTable INNER JOIN " _
          & " GlacTable ON ComnTable.CODEFLABORACCT = GlacTable.GLACCTREF"
   bSqlRows = GetDataSet(RdoLab)
   If bSqlRows Then
      With RdoLab
         sAcct = "" & Trim(!GLACCTNO)
         sDesc = "" & Trim(!GLDESCR)
      End With
   End If
   Set RdoLab = Nothing
   Exit Sub
   
   DiaErr1:
   sProcName = "getlabor"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub PrintReport()
   Dim sTemp As String
   Dim sType As String
   Dim sLaborAcct As String
   Dim sLaborDesc As String
   Dim sJournal As String
   
   MouseCursor 13
   cmbJrn.SetFocus
   
   sJournal = Trim(cmbJrn)
   sType = UCase(Left(cmbJrn, 2))
   
   optPrn.Enabled = False
   optDis.Enabled = False
   On Error GoTo DiaErr1
   SetMdiReportsize MdiSect
   
   MdiSect.crw.Formulas(0) = "CompanyName='" & sFacility & "'"
   MdiSect.crw.Formulas(1) = "RequestBy='Requested By: " _
                        & Secure.UserInitials & "'"
   
   sSql = ""
   
   Select Case sType
      Case "SJ"
         MdiSect.crw.Formulas(2) = "Journal1='Sales Journal Number " & Right(cmbJrn, Len(cmbJrn) - 3) & "'"
         MdiSect.crw.ReportFileName = sReportPath & "finjr01sj.rpt"
         sSql = "{JritTable.DCHEAD}='" & Trim(cmbJrn) & "'"
      Case "PJ"
         MdiSect.crw.Formulas(2) = "Journal1='Purchases Journal Number " & Right(cmbJrn, Len(cmbJrn) - 3) & "'"
         MdiSect.crw.ReportFileName = sReportPath & "finjr01pj.rpt"
         sSql = "{JritTable.DCHEAD}='" & Trim(cmbJrn) & "' AND {JritTable.DCCREDIT} = 0"
      Case "CC"
         MdiSect.crw.Formulas(2) = "Journal1='Computer Check Journal Number " & Right(cmbJrn, Len(cmbJrn) - 3) & "'"
         MdiSect.crw.ReportFileName = sReportPath & "finjr01cc.rpt"
         sSql = "{JritTable.DCHEAD}='" & Trim(cmbJrn) & "'"
      Case "XC"
         MdiSect.crw.Formulas(2) = "Journal1='External Check Journal Number " & Right(cmbJrn, Len(cmbJrn) - 3) & "'"
         MdiSect.crw.ReportFileName = sReportPath & "finjr01xc.rpt"
         sSql = "{JritTable.DCHEAD}='" & Trim(cmbJrn) & "'"
      Case "CR"
         MdiSect.crw.Formulas(2) = "Journal1='Cash Receipts Journal Number " & Right(cmbJrn, Len(cmbJrn) - 3) & "'"
         MdiSect.crw.ReportFileName = sReportPath & "finjr01cr.rpt"
         sSql = "{JritTable.DCHEAD}='" & Trim(cmbJrn) & "'"
         MdiSect.crw.Formulas(7) = "DCHEAD='" & sJournal & "'"
      Case "TJ"
         GetLabor sLaborAcct, sLaborDesc
         MdiSect.crw.Formulas(2) = "Title1='Time Charges Journal'"
         MdiSect.crw.ReportFileName = sReportPath & "finjr01tj.rpt"
         sSql = "{TcitTable.TCGLJOURNAL} = '" & Trim(cmbJrn) & "'"
   End Select
   
   If sType <> "TJ" Then
      MdiSect.crw.Formulas(3) = "Journal2='" & lblDsc & "'"
      MdiSect.crw.Formulas(4) = "Journal3='From " & lblStart & " Thru " & lblEnd & "'"
      MdiSect.crw.Formulas(5) = "Journal4='Type: " & sType & " Journal Open " & lblOpen & " Journal Closed " & lblClose & "'"
      If Trim(lblClose) = "" Then
         MdiSect.crw.Formulas(6) = "Journal5='********** This Journal Is Still Open!  Do Not Use This Report For Posting To General Ledger! **********'"
      End If
   Else
      MdiSect.crw.Formulas(3) = "DefaultLaborAcct='" & sLaborAcct & "'"
      MdiSect.crw.Formulas(4) = "LaborDesc='" & sLaborDesc & "'"
   End If
   
   MdiSect.crw.SelectionFormula = sSql
   
   SetCrystalAction Me
   
   MouseCursor 0
   optPrn.Enabled = True
   optDis.Enabled = True
   
   Exit Sub
   
   DiaErr1:
   optPrn.Enabled = True
   optDis.Enabled = True
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Public Sub SaveOptions()
   SaveSetting "Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter
End Sub

Public Sub GetOptions()
   lblPrinter = GetSetting("Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter)
   If lblPrinter = "" Then
      lblPrinter = TTDEFAULT
   End If
End Sub

Private Sub lblDsc_Change()
   If Left(lblDsc, 6) = "*** No" Then
      lblDsc.ForeColor = ES_RED
   Else
      lblDsc.ForeColor = vbBlack
   End If
End Sub

Private Sub optDis_Click()
   PrintReport
End Sub

Private Sub optPrn_Click()
   PrintReport
End Sub

Public Sub FillJournals()
   Dim rdoJrn As rdoResultset
   Dim sType As String
   
   On Error GoTo DiaErr1
   
   cmbJrn.Clear
   
   If cmbTyp = "ALL" Then
      sType = " AND MJTYPE IN ('SJ','PJ','XC','CC','CR','TJ')"
   Else
      sType = " AND MJTYPE = '" & cmbTyp & "'"
   End If
   
   sSql = "SElECT MJGLJRNL FROM JrhdTable WHERE MJFY=" _
          & Trim(cmbFyr) & sType
   bSqlRows = GetDataSet(rdoJrn, ES_FORWARD)
   If bSqlRows Then
      With rdoJrn
         Do Until .EOF
            If Not IsNull(.rdoColumns(0)) Then _
                          AddComboStr cmbJrn.hWnd, "" & Trim(.rdoColumns(0))
            .MoveNext
         Loop
         .Cancel
      End With
      If cmbJrn.ListCount > 0 Then
         cmbJrn = cmbJrn.List(0)
         GetJrnId
      End If
   Else
      lblNum = ""
      lblOpen = ""
      lblClose = ""
      lblDsc = "*** No Journals Found In Fy ***"
   End If
   Exit Sub
   
   DiaErr1:
   sProcName = "filljourn"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Public Function GetJrnId() As Byte
   Dim RdoJid As rdoResultset
   
   On Error GoTo DiaErr1
   
   sSql = "SELECT MJTYPE,MJFY,MJNO,MJDESCRIPTION,MJGLJRNL," _
          & "MJOPENED,MJCLOSED,MJSTART,MJEND FROM JrhdTable WHERE (MJFY=" _
          & cmbFyr & " AND MJGLJRNL='" & cmbJrn & "')"
   bSqlRows = GetDataSet(RdoJid)
   If bSqlRows Then
      With RdoJid
         lblDsc = "" & Trim(!MJDESCRIPTION)
         lblNum = Format(!MJNO, "####0")
         lblOpen = "" & Format(!MJOPENED, "mm/dd/yy")
         lblClose = "" & Format(!MJCLOSED, "mm/dd/yy")
         lblStart = "" & Format(!MJSTART, "mm/dd/yy")
         lblEnd = "" & Format(!MJEND, "mm/dd/yy")
         .Cancel
      End With
      GetJrnId = 1
   Else
      lblNum = ""
      lblOpen = ""
      lblClose = ""
      lblDsc = "*** No Journal Found ***"
      GetJrnId = 0
   End If
   Exit Function
   
   DiaErr1:
   sProcName = "getjrnid"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Function

Private Sub ShowPrinters_Click(Value As Integer)
   SysPrinters.Show
   ShowPrinters.Value = False
End Sub
