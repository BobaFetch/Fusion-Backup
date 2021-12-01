VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form diaAPf06a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Export AP Invoice Activity To QuickBooks ® IIF"
   ClientHeight    =   5235
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
   ScaleHeight     =   5235
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdPrevious 
      Height          =   320
      Left            =   2520
      Picture         =   "diaAPf06a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   25
      TabStop         =   0   'False
      ToolTipText     =   "Show Exported Invoices"
      Top             =   1560
      Width           =   350
   End
   Begin VB.Frame Frame1 
      Caption         =   "Export Mode"
      Height          =   795
      Left            =   120
      TabIndex        =   24
      Top             =   2340
      Width           =   3735
      Begin VB.OptionButton optExportAll 
         Caption         =   "Export All Invoices in date range"
         Height          =   195
         Left            =   60
         TabIndex        =   3
         ToolTipText     =   "This option exports all invoices in the date range including those you have previously exported"
         Top             =   480
         Width           =   3555
      End
      Begin VB.OptionButton optExportNew 
         Caption         =   "Export only New invoices in date range"
         Height          =   195
         Left            =   60
         TabIndex        =   2
         ToolTipText     =   "This option only exports invoices you have not exported before"
         Top             =   240
         Value           =   -1  'True
         Width           =   3555
      End
   End
   Begin VB.TextBox txtFreightAccount 
      Height          =   285
      Left            =   1320
      TabIndex        =   6
      Top             =   3960
      Width           =   2535
   End
   Begin VB.CommandButton cmdVew 
      Height          =   320
      Left            =   2520
      Picture         =   "diaAPf06a.frx":04DA
      Style           =   1  'Graphical
      TabIndex        =   20
      TabStop         =   0   'False
      ToolTipText     =   "Show Selected Invoices"
      Top             =   1200
      Width           =   350
   End
   Begin VB.TextBox txtTaxAccount 
      Height          =   285
      Left            =   1320
      TabIndex        =   5
      Top             =   3600
      Width           =   2535
   End
   Begin VB.TextBox txtAPAccount 
      Height          =   285
      Left            =   1320
      TabIndex        =   4
      Top             =   3240
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
      FormDesignHeight=   5235
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
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Browse For File Location"
      Top             =   4320
      Width           =   375
   End
   Begin VB.TextBox txtPath 
      Height          =   285
      Left            =   1320
      TabIndex        =   7
      Top             =   4320
      Width           =   3735
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Export"
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
      Left            =   4680
      TabIndex        =   9
      ToolTipText     =   "Build QuickBooks Export"
      Top             =   4800
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
      TabIndex        =   14
      Top             =   4800
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
      TabIndex        =   10
      TabStop         =   0   'False
      ToolTipText     =   "Save And Exit"
      Top             =   120
      Width           =   875
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
      Autosize        =   2
      RoundedCorners  =   0   'False
      BevelWidth      =   0
      Outline         =   0   'False
      PictureUp       =   "diaAPf06a.frx":09B4
      PictureDn       =   "diaAPf06a.frx":0AFA
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Previously Exported"
      Height          =   285
      Index           =   7
      Left            =   120
      TabIndex        =   27
      Top             =   1560
      Width           =   1545
   End
   Begin VB.Label lblPrevious 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1800
      TabIndex        =   26
      Tag             =   "1"
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Just Exported"
      Height          =   285
      Index           =   5
      Left            =   120
      TabIndex        =   23
      Top             =   1920
      Width           =   1305
   End
   Begin VB.Label lblExported 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1800
      TabIndex        =   22
      Tag             =   "1"
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Freight Account"
      Height          =   285
      Index           =   12
      Left            =   120
      TabIndex        =   21
      Top             =   3960
      Width           =   1905
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Account"
      Height          =   285
      Index           =   6
      Left            =   120
      TabIndex        =   19
      Top             =   3600
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "AP Account"
      Height          =   285
      Index           =   4
      Left            =   120
      TabIndex        =   18
      Top             =   3240
      Width           =   1305
   End
   Begin VB.Label lblFound 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1800
      TabIndex        =   17
      Tag             =   "1"
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Invoices Selected"
      Height          =   285
      Index           =   3
      Left            =   120
      TabIndex        =   16
      Top             =   1200
      Width           =   1305
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Disk Path\File"
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   15
      Top             =   4320
      Width           =   1185
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "End Date"
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   13
      Top             =   720
      Width           =   825
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Start Date"
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   12
      Top             =   360
      Width           =   825
   End
End
Attribute VB_Name = "diaAPf06a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of                    ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***

Option Explicit

'*************************************************************************************
' diaAPf06a - Export AP Activity To QuickBooks
'
' Notes:
'
' Created: 04/04/05 (TEL)
'*************************************************************************************

Dim bOnLoad As Byte
Dim sInvHdr(14) As String
Dim sInvItHdr(14) As String
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

Private Sub cmdPrevious_Click()
   Dim RdoInv As ADODB.Recordset
   Dim sItem As String
   
   'make sure dates are valid
   If Not IsValidDate(txtstart) Or Not IsValidDate(txtEnd) Then
      MsgBox "Invalid date or date range"
      Exit Sub
   End If
   
   '    VewAPInv.lblCaption = "Invoices Selected For Export"
   '    sSql = "SELECT VINO,VIDATE,VIDTRECD,VENICKNAME,VIDUE " _
   '        & "FROM VihdTable inv INNER JOIN VndrTable vdr ON inv.VIVENDOR = vdr.VEREF " _
   '        & "WHERE VIDTRECD >= '" & txtStart & "' AND VIDTRECD <='" & txtEnd & "' " _
   '        & "AND VIDUE > 0 " & vbCrLf
   '
   VewAPInv.lblCaption = "Invoices Selected For Export"
   sSql = "SELECT VINO,VIDATE,VIDTRECD,VIVENDOR,VIDUE " _
          & "FROM VihdTable vi " & vbCrLf _
          & "join QbapTable qt on vi.VIVENDOR = qt.VENDOR " & vbCrLf _
          & "and vi.VINO = qt.INVOICE " & vbCrLf _
          & "WHERE VIDTRECD >= '" & txtstart & "' AND VIDTRECD <='" & txtEnd & "' " _
          & "AND VIDUE > 0 " & vbCrLf _
          & "ORDER BY VIVENDOR, VINO"
   
   Debug.Print sSql
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoInv)
   If bSqlRows Then
      With RdoInv
         While Not .EOF
            Dim sNo, sAmt, sName As String
            sNo = Trim(!VINO)
            sAmt = Format("" & Trim(!VIDUE), "#,###,##0.00")
            sName = Trim(!VIVENDOR)
            sItem = _
                    sName & Space(12 - Len(sName)) _
                    & sNo & Space(22 - Len(sNo)) _
                    & Format(Trim(!VIDATE), "mm/dd/yy") & Space(2) _
                    & Format(Trim(!VIDTRECD), "mm/dd/yy") & Space(2) _
                    & Space(12 - Len(sAmt)) & sAmt
            VewAPInv.lstInv.AddItem sItem
            .MoveNext
         Wend
      End With
   End If
   
   Set RdoInv = Nothing
   VewAPInv.Show
End Sub

Private Sub cmdVew_Click()
   Dim RdoInv As ADODB.Recordset
   Dim sItem As String
   
   'make sure dates are valid
   If Not IsValidDate(txtstart) Or Not IsValidDate(txtEnd) Then
      MsgBox "Invalid date or date range"
      Exit Sub
   End If
   
   '    VewAPInv.lblCaption = "Invoices Selected For Export"
   '    sSql = "SELECT VINO,VIDATE,VIDTRECD,VENICKNAME,VIDUE " _
   '        & "FROM VihdTable inv INNER JOIN VndrTable vdr ON inv.VIVENDOR = vdr.VEREF " _
   '        & "WHERE VIDTRECD >= '" & txtStart & "' AND VIDTRECD <='" & txtEnd & "' " _
   '        & "AND VIDUE > 0 " & vbCrLf
   '
   VewAPInv.lblCaption = "Invoices Selected For Export"
   sSql = "SELECT VINO,VIDATE,VIDTRECD,VIVENDOR,VIDUE " _
          & "FROM VihdTable " & vbCrLf _
          & "WHERE VIDTRECD >= '" & txtstart & "' AND VIDTRECD <='" & txtEnd & "' " _
          & "AND VIDUE > 0 " & vbCrLf
   
   
   'if new only, add clause to exclude previously exported invoices
   If optExportNew.Value Then
      sSql = sSql & "and VINO not in " & vbCrLf _
             & "( select INVOICE from QbapTable where VENDOR = VIVENDOR and INVOICE = VINO )" & vbCrLf
   End If
   sSql = sSql & "order by VIVENDOR, VINO" & vbCrLf
   
   Debug.Print sSql
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoInv)
   If bSqlRows Then
      With RdoInv
         While Not .EOF
            Dim sNo, sAmt, sName As String
            sNo = Trim(!VINO)
            sAmt = Format("" & Trim(!VIDUE), "#,###,##0.00")
            sName = Trim(!VIVENDOR)
            sItem = _
                    sName & Space(12 - Len(sName)) _
                    & sNo & Space(22 - Len(sNo)) _
                    & Format(Trim(!VIDATE), "mm/dd/yy") & Space(2) _
                    & Format(Trim(!VIDTRECD), "mm/dd/yy") & Space(2) _
                    & Space(12 - Len(sAmt)) & sAmt
            VewAPInv.lstInv.AddItem sItem
            .MoveNext
         Wend
      End With
   End If
   
   Set RdoInv = Nothing
   VewAPInv.Show
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad = 1 Then
      bOnLoad = 0
      'IniQB
      GetInvoices
   End If
   
   MouseCursor 0
End Sub

Private Sub Form_Load()
   Dim rdoStr As ADODB.Recordset
   
   FormLoad Me, ES_DONTLIST
   FormatControls
   bOnLoad = 1
   
   'make sure QbapTable, which contains exported vendor nickname & invoice # for each
   'exported invoice, exists.  if it doesn't, create it
   On Error Resume Next
   sSql = "create table QbapTable " _
          & "( VENDOR varchar(10), INVOICE varchar(20), WHENEXPORTED datetime default getdate())"
   clsADOCon.ExecuteSQL sSql
   
   'get default account values for AP export
   Dim sAPAccount As String, sFreightAccount As String, sTaxAccount As String
   
   sSql = "select COAPACCT, COPJTAXACCT, COPJTFRTACCT from ComnTable"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoStr)
   If bSqlRows Then
      sAPAccount = Trim(rdoStr("COAPACCT"))
      sFreightAccount = Trim(rdoStr("COPJTAXACCT"))
      sTaxAccount = Trim(rdoStr("COPJTFRTACCT"))
   End If
   Set rdoStr = Nothing
   'retrieve settings used last time
   txtstart = GetSetting("Esi2000", "Quickbooks", "StartDate", Date)
   txtEnd = GetSetting("Esi2000", "Quickbooks", "EndDate", Date)
   txtAPAccount = GetSetting("Esi2000", "Quickbooks", "APAccount", sAPAccount)
   txtFreightAccount = GetSetting("Esi2000", "Quickbooks", "APFreightAccount", sFreightAccount)
   txtTaxAccount = GetSetting("Esi2000", "Quickbooks", "APTaxAccount", sTaxAccount)
   Me.txtPath = GetSetting("Esi2000", "Quickbooks", "APPath", "")
   
   'Columns for QuickBooks Bill Heading
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
   sInvHdr(12) = "DUEDATE"
   sInvHdr(13) = "TERMS"
   
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
   sInvItHdr(12) = "SERVICEDATE"
   sInvItHdr(13) = "OTHER2"
   
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
   Set diaAPf06a = Nothing
End Sub

Public Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Private Sub optExportAll_Click()
   GetInvoices
End Sub

Private Sub optExportNew_Click()
   GetInvoices
End Sub

Private Sub txtEnd_Change()
   GetInvoices
End Sub

Private Sub txtend_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtEnd_KeyDown(KeyCode As Integer, Shift As Integer)
   Debug.Print KeyCode
   On Error GoTo whoops
   If KeyCode = vbKeyUp Then 'up
      txtEnd.Text = DateAdd("d", 1, txtEnd.Text)
   ElseIf KeyCode = vbKeyDown Then 'down
      txtEnd.Text = DateAdd("d", -1, txtEnd.Text)
   End If
   Exit Sub
whoops:
End Sub

Private Sub txtEnd_LostFocus()
   txtEnd = CheckDate(txtEnd)
   'GetInvoices
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

Private Sub txtStart_Change()
   GetInvoices
End Sub

Private Sub txtstart_DropDown()
   ShowCalendar Me
End Sub

Private Sub GetInvoices()
   
   'get count of invoices selected
   Dim RdoInv As ADODB.Recordset
   
   If Not IsValidDate(txtstart) Or Not IsValidDate(txtEnd) Then
      lblFound = ""
      Exit Sub
   End If
   
   'count the selected invoices in the database between these dates
   'don't count credit memos ( < 0 amount )
   On Error GoTo DiaErr1
   sSql = "select count(VINO) from VihdTable where VIDTRECD >= '" & txtstart & "'" & vbCrLf _
          & "and VIDTRECD <= '" & txtEnd & "' and VIDUE > 0" & vbCrLf
   
   'if getting only new invoices, add exclusion to the where clause
   If optExportNew.Value Then
      sSql = sSql & "and VINO not in " & vbCrLf _
             & "( select VINO from QbapTable where VENDOR = VIVENDOR and INVOICE = VINO )" & vbCrLf
   End If
   Debug.Print sSql
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoInv, ES_FORWARD)
   lblFound = ""
   cmdGo.enabled = False
   If bSqlRows Then
      If RdoInv.Fields(0) > 0 Then
         lblFound = RdoInv.Fields(0)
         cmdGo.enabled = True
      End If
   End If
   Set RdoInv = Nothing
   
   'now do a count of invoices exported for this date range
   sSql = "select count(VINO) from VihdTable vi " & vbCrLf _
          & "join QbapTable qt on vi.VIVENDOR = qt.VENDOR" & vbCrLf _
          & "and vi.VINO = qt.INVOICE" & vbCrLf _
          & "where VIDTRECD >= '" & txtstart & "'" & vbCrLf _
          & "and VIDTRECD <= '" & txtEnd & "' and VIDUE > 0" & vbCrLf
   
   Debug.Print sSql
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoInv, ES_FORWARD)
   lblPrevious = ""
   If bSqlRows Then
      If RdoInv.Fields(0) > 0 Then
         lblPrevious = RdoInv.Fields(0)
      End If
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
   Dim RdoInv As ADODB.Recordset
   
   'verify that all fields have valid data
   '    On Error Resume Next
   '    Dim v As Variant
   '    v = DateValue(txtStart)
   '    v = DateValue(txtEnd)
   '    If Err Then
   '        MsgBox "Invalid date or date range"
   '        Exit Sub
   '    End If
   
   lblExported = ""
   
   If Not IsValidDate(txtstart) Or Not IsValidDate(txtEnd) Then
      MsgBox "Invalid date or date range"
      Exit Sub
   End If
   
   If Trim(txtAPAccount) = "" Or Trim(txtTaxAccount) = "" Or Trim(txtFreightAccount) = "" Then
      MsgBox "Missing account information"
      Exit Sub
   End If
   
   If CInt("" & lblFound) = 0 Then
      MsgBox "No invoices in this date range"
      Exit Sub
   End If
   
   If Trim(txtPath) = "" Then
      MsgBox "You must specify disk file/path"
      Exit Sub
   Else
      Select Case MsgBox("This file already exists.  Do you want to overwrite it?", _
                         vbYesNoCancel, "ESI QuickBooks AP Export")
         Case vbYes
         Case Else
            Exit Sub
      End Select
   End If
   
   '    Select Case MsgBox("OK to proceed?", vbYesNoCancel, "ESI QuickBooks AP Export")
   '    Case vbYes
   '    Case Else
   '        Exit Sub
   '    End Select
   '
   iFile = FreeFile
   'Trim (txtPath)
   Open Trim(txtPath) For Output As iFile
   If Err Then
      MsgBox "Unable to open file " & Trim(txtPath)
      Exit Sub
   End If
   
   Dim i As Integer
   Dim sMsg As String
   Dim sbuf As String
   On Error GoTo DiaErr1
   MouseCursor 13
   
   'get general account information
   Dim sApAcct As String, sTaxAcct As String, sFrtAcct As String
   sApAcct = Trim(txtAPAccount)
   sTaxAcct = Trim(txtTaxAccount)
   sFrtAcct = Trim(txtFreightAccount)
   
   'read all data as strings
   sSql = "select VIVENDOR as VENDOR, VINO as DOCNUM, VITACCOUNT as ACCT, " & vbCrLf _
          & "cast(cast(sum(VITCOST*VITQTY + VITADDERS) as decimal(12,2)) as varchar(12) ) as ACCTAMT, " & vbCrLf _
          & "convert( char(10), max(VIDTRECD), 101) as POSTDATE, " & vbCrLf _
          & "cast( max(VIFREIGHT) as varchar(10) ) as FREIGHT, " & vbCrLf _
          & "cast( max(VITAX) as varchar(12) ) as TAX, " & vbCrLf _
          & "cast(cast(max(VIDUE) as decimal(12,2)) as varchar(12) ) as INVAMT, " _
          & "convert( char(10), max(VIDUEDATE), 101) as DUEDATE " & vbCrLf _
          & "from VihdTable inv " & vbCrLf _
          & "join ViitTable item on item.VITNO = inv.VINO " & vbCrLf _
          & "and item.VITVENDOR = inv.VIVENDOR " & vbCrLf _
          & "where VIDTRECD >= '" & txtstart & "' and VIDTRECD <='" & txtEnd & "' " & vbCrLf _
          & "and VIDUE > 0 " & vbCrLf
   
   'if new only, add clause to exclude previously exported invoices
   If optExportNew.Value Then
      sSql = sSql & "and VINO not in " & vbCrLf _
             & "( select VINO from QbapTable where VENDOR = inv.VIVENDOR and INVOICE = inv.VINO)" & vbCrLf
   End If
   
   'add group by clause at end
   sSql = sSql & "group by VIVENDOR, VINO, VITACCOUNT "
   Debug.Print sSql
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoInv, ES_STATIC)
   
   If bSqlRows Then
      prg1.Visible = True
      prg1.max = Val(lblFound)
      'Open sFileName For Output As iFile
      
      With RdoInv
         
         'output header line 1: !TRNS...
         sbuf = ""
         For i = 0 To UBound(sInvHdr)
            sbuf = sbuf & sInvHdr(i) & vbTab
         Next
         Print #iFile, sbuf
         
         'output header line 2: !SPL
         sbuf = ""
         For i = 0 To UBound(sInvItHdr)
            sbuf = sbuf & sInvItHdr(i) & vbTab
         Next
         Print #iFile, sbuf
         
         'output header line 3: !ENDTRNS
         Print #iFile, sQBEndTrans
         
         'now loop through invoice subtotals for each acct and output a line for each
         'fields are:
         '   VENDOR      vendor nickname
         '   DOCNUM      invoice number
         '   ACCT        account number
         '   ACCTAMT     account debit amount
         '   POSTDATE    date posted
         '   FREIGHT     invoice freight amount
         '   TAX         invoice tax amount
         '   INVAMT      invoice total amount
         '   DUEDATE     invoice due date
         
         While Not .EOF
            Dim sName As String, sNo As String, sAcct As String, sAmt As String
            Dim sDate As String, sFrt As String, sTax As String, sTotal As String, sDueDate As String
            Dim sPriorName As String, sPriorNo As String, sPriorDate As String
            Dim sPriorFrtAmt As String, sPriorTaxAmt As String
            
            sName = Trim(!VENDOR)
            sNo = Trim(!DOCNUM)
            sAcct = Trim(!acct)
            sAmt = Trim(!ACCTAMT)
            sDate = Trim(!postDate)
            sFrt = Trim(!FREIGHT)
            sTax = Trim(!tax)
            sTotal = Trim(!INVAMT)
            sDueDate = Trim(!DUEDATE)
            
            Debug.Print sName & " " & sNo & " " & sAcct & " " & sAmt & " " _
               ; sDate & " " & sFrt & " " & sTax & " " & sTotal & " " & sDueDate
            
            'if new invoice, output prior invoice tax and freight and then current invoice header
            If sName <> sPriorName Or sNo <> sPriorNo Then
               If sPriorName <> "" Then
                  OutputAccount iFile, sPriorDate, sPriorNo, "Freight", sFrtAcct, sPriorFrtAmt
                  OutputAccount iFile, sPriorDate, sPriorNo, "Tax", sTaxAcct, sPriorTaxAmt
                  
                  'output end of transaction
                  Print #iFile, Right(sQBEndTrans, Len(sQBEndTrans) - 1)
                  prg1.Value = prg1.Value + 1
                  
               End If
               
               'output the total for the invoice
               OutputInvoice iFile, sDate, sNo, "", sApAcct, sName, "-" & sTotal, sDueDate, ""
               lblExported = CStr(CInt("0" + lblExported) + 1)
               
            End If
            
            'now output current account total
            OutputAccount iFile, sDate, sNo, "", sAcct, sAmt
            
            'retain prior values
            sPriorName = sName
            sPriorNo = sNo
            sPriorDate = sDate
            sPriorFrtAmt = sFrt
            sPriorTaxAmt = sTax
            .MoveNext
         Wend
      End With
      
      'output frt and tax for the last invoice
      If sPriorName <> "" Then
         OutputAccount iFile, sPriorDate, sPriorNo, "Freight", sFrtAcct, sPriorFrtAmt
         OutputAccount iFile, sPriorDate, sPriorNo, "Tax", sTaxAcct, sPriorTaxAmt
         
         'output end of the last transaction
         Print #iFile, Right(sQBEndTrans, Len(sQBEndTrans) - 1)
         prg1.Value = prg1.Value + 1
      End If
      
      'save successful settings
      SaveSetting "Esi2000", "Quickbooks", "StartDate", txtstart
      SaveSetting "Esi2000", "Quickbooks", "EndDate", txtEnd
      SaveSetting "Esi2000", "Quickbooks", "APAccount", txtAPAccount
      SaveSetting "Esi2000", "Quickbooks", "APFreightAccount", txtFreightAccount
      SaveSetting "Esi2000", "Quickbooks", "APTaxAccount", txtTaxAccount
      SaveSetting "Esi2000", "Quickbooks", "APPath", txtPath
      
      MsgBox "Quickbooks AP IIF File created with " & lblExported _
         & " invoices.  You can view or modify the file in Excel before importing it into QuickBooks."
      GetInvoices
   Else
      sMsg = "Could Not Build QuickBooks" & vbCrLf & "AP Export."
      MsgBox sMsg, vbExclamation, Caption
   End If
   
   
CleanUp:
   prg1.Value = 0
   prg1.Visible = False
   Close iFile
   Set RdoInv = Nothing
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "BuildQBExport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   GoTo CleanUp
End Sub

Private Sub OutputAccount(iFile As Integer, sDate As String, sInvNo As String, _
                          sMemo As String, sAcct As String, sAmt As String)
   'output account detail line item
   'if any field is blank, no output occurs
   
   Dim sbuf As String
   
   If sDate = "" Or sInvNo = "" Or sAcct = "" Or sAmt = "" Or sAmt = "0" Then
      Exit Sub
   End If
   
   sbuf = Right(sInvItHdr(0), Len(sInvItHdr(0)) - 1) & vbTab
   sbuf = sbuf & vbTab
   sbuf = sbuf & "BILL" & vbTab
   sbuf = sbuf & sDate & vbTab
   sbuf = sbuf & sAcct & vbTab
   sbuf = sbuf & vbTab
   sbuf = sbuf & vbTab
   sbuf = sbuf & sAmt & vbTab
   sbuf = sbuf & sInvNo & vbTab
   sbuf = sbuf & sMemo & vbTab
   sbuf = sbuf & vbTab
   sbuf = sbuf & vbTab
   sbuf = sbuf & vbTab
   Print #iFile, sbuf
End Sub

Private Sub OutputInvoice(iFile As Integer, sDate As String, sInvNo As String, _
                          sMemo As String, sAcct As String, _
                          sName As String, sAmt As String, sDueDate As String, sTerms As String)
   'output invoice header line
   Dim sbuf As String
   
   'sbuf = Right(sInvHdr(0), Len(sInvItHdr(0)) - 1) & vbTab
   sbuf = Mid(sInvHdr(0), 2) & vbTab
   sbuf = sbuf & vbTab
   sbuf = sbuf & "BILL" & vbTab
   sbuf = sbuf & sDate & vbTab
   sbuf = sbuf & sAcct & vbTab
   sbuf = sbuf & sName & vbTab
   sbuf = sbuf & vbTab
   sbuf = sbuf & sAmt & vbTab
   sbuf = sbuf & sInvNo & vbTab
   sbuf = sbuf & sMemo & vbTab
   sbuf = sbuf & vbTab
   sbuf = sbuf & vbTab
   sbuf = sbuf & sDueDate & vbTab
   sbuf = sbuf & sTerms & vbTab
   Print #iFile, sbuf
   
   'output the exported invoice # to the QbapTable Table
   Dim rdo As ADODB.Recordset
   sSql = "declare @ct int" & vbCrLf _
          & "select @ct = count(*) from QbapTable where VENDOR = '" & sName & "' " & vbCrLf _
          & "and INVOICE = '" & sInvNo & "'" & vbCrLf _
          & "if @ct = 0" & vbCrLf _
          & "    insert QbapTable ( VENDOR, INVOICE ) values ( '" _
          & sName & "', '" & sInvNo & "')"
   Debug.Print sSql
   clsADOCon.ExecuteSQL sSql
End Sub

Private Sub txtStart_KeyDown(KeyCode As Integer, Shift As Integer)
   Debug.Print KeyCode
   On Error GoTo whoops
   If KeyCode = vbKeyUp Then 'up
      txtstart.Text = DateAdd("d", 1, txtstart.Text)
   ElseIf KeyCode = vbKeyDown Then 'down
      txtstart.Text = DateAdd("d", -1, txtstart.Text)
   End If
   Exit Sub
whoops:
End Sub

Private Sub txtstart_LostFocus()
   txtstart = CheckDate(txtstart)
End Sub

Private Function IsValidDate(dt As String)
   IsValidDate = False
   If Not IsDate(dt) Then Exit Function
   Dim n As Integer
   n = InStr(1, dt, "/")
   If n = 0 Then Exit Function
   n = InStr(n + 1, dt, "/")
   If n = 0 Then Exit Function
   IsValidDate = True
End Function
