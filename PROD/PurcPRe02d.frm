VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form PurcPRe02d 
   BackColor       =   &H80000018&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Service Part Numbers"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6210
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   6210
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5040
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   975
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6600
      Top             =   4200
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4065
      FormDesignWidth =   6210
   End
   Begin MSFlexGridLib.MSFlexGrid Grd 
      Height          =   2655
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Click Or Select And Press Enter"
      Top             =   1080
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   4683
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      HighLight       =   2
      ScrollBars      =   2
   End
   Begin VB.Label Z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Remaining Quantity"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   7
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label lblRem 
      BackStyle       =   0  'Transparent
      Caption         =   "Qty Rem"
      Height          =   255
      Left            =   2760
      TabIndex        =   6
      Top             =   600
      Width           =   1095
   End
   Begin VB.Image Chkyes 
      Height          =   210
      Left            =   120
      Picture         =   "PurcPRe02d.frx":0000
      Top             =   2760
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Chkno 
      Height          =   210
      Left            =   360
      Picture         =   "PurcPRe02d.frx":038A
      Top             =   2760
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Label Z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Run"
      Height          =   255
      Index           =   8
      Left            =   2760
      TabIndex        =   5
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Z1 
      BackStyle       =   0  'Transparent
      Caption         =   "MO"
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label lblRun 
      BackStyle       =   0  'Transparent
      Caption         =   "Mo Run"
      Height          =   255
      Left            =   3600
      TabIndex        =   3
      Top             =   240
      Width           =   615
   End
   Begin VB.Label lblMon 
      BackStyle       =   0  'Transparent
      Caption         =   "Mo Number"
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "PurcPRe02d"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007)) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
'12/15/03 New (requested by LH)
'1/31/07 fixed grid and UpdateItem 7.2.3
Option Explicit
Dim AdoQry As ADODB.Command
Dim AdoParameter1 As ADODB.Parameter
Dim ADOParameter2 As ADODB.Parameter
Dim bOnLoad As Byte
Dim bItemSelected As Byte

Dim sServPart(100, 4) As String
'0 = OPNO
'1 = PARTREF
'2 = PADESC
'3 = PURCHASED
'4 = Unit Price

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmdCan_Click()
   Dim bResponse As Byte
   Dim sMsg As String
   If bItemSelected = 0 Then
      sMsg = "You Have Not Properly Selected An Item " & vbCr _
             & "(Double Click The Item Or Scroll And Press Enter)." & vbCr _
             & "Continue To Close Anyway?"
      bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
      If bResponse = vbYes Then Unload Me
   Else
      Unload Me
   End If
   
End Sub



Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      lblRem = Format(GetMoQty(), "######0.000")
      FillGrid
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Deactivate()
   PurcPRe02c.optSel.Value = vbUnchecked
   Unload Me
   
End Sub


Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST
   Move PurcPRe02c.Left + 1000, PurcPRe02c.Top + 1000
   BackColor = Es_HelpBackGroundColor
   FormatControls
   sSql = "select OPREF,OPRUN,OPNO,OPSERVPART,OPPURCHASED,OPSVCUNIT," _
          & "PARTREF,PARTNUM,PADESC FROM RnopTable,PartTable " _
          & "where OPSERVPART=PARTREF AND (OPREF= ? " _
          & "AND OPRUN= ? AND OPSERVPART<>'') ORDER BY OPNO"
   Set AdoQry = New ADODB.Command
   AdoQry.CommandText = sSql
   
   Set AdoParameter1 = New ADODB.Parameter
   AdoParameter1.Type = adChar
   AdoParameter1.SIZE = 30
   
   Set ADOParameter2 = New ADODB.Parameter
   ADOParameter2.Type = adInteger
   
   AdoQry.Parameters.Append AdoParameter1
   AdoQry.Parameters.Append ADOParameter2
   
   With Grd
      .ColAlignment(0) = 0
      .ColAlignment(1) = 0
      .ColAlignment(2) = 0
      .row = 0
      .Col = 0
      .Text = "Op No"
      .ColWidth(0) = 650
      .Col = 1
      .Text = "Service Part Number"
      .ColWidth(1) = 2500
      .Col = 2
      .Text = "Description"
      .ColWidth(2) = 2050
      .Col = 3
      .Text = "Purchased"
      .ColWidth(3) = 650
      .Col = 0
   End With
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   MouseCursor 0
   Set AdoParameter1 = Nothing
   Set ADOParameter2 = Nothing
   Set AdoQry = Nothing
   Set PurcPRe02d = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub



Private Sub FillGrid()
   Dim RdoRte As ADODB.Recordset
   Dim iRows As Integer
   
   AdoQry.Parameters(0).Value = Compress(PurcPRe02c.cmbMon)
   AdoQry.Parameters(1).Value = Val(PurcPRe02c.cmbRun)
   bSqlRows = clsADOCon.GetQuerySet(RdoRte, AdoQry, ES_FORWARD)
   If bSqlRows Then
      With RdoRte
         Do Until .EOF
            iRows = iRows + 1
            If iRows > 1 Then Grd.Rows = Grd.Rows + 1
            Grd.row = iRows
            Grd.Col = 0
            Grd.Text = Format(!opNo, "000")
            sServPart(Grd.row, 0) = Grd.Text
            Grd.Col = 1
            Grd.Text = "" & Trim(!PartNum)
            sServPart(Grd.row, 1) = Grd.Text
            Grd.Col = 2
            Grd.Text = "" & Trim(!PADESC)
            sServPart(Grd.row, 2) = Grd.Text
            Grd.Col = 3
            If !OPPURCHASED = 1 Then
               Set Grd.CellPicture = Chkyes.Picture
               sServPart(Grd.row, 3) = "1"
            Else
               Set Grd.CellPicture = Chkno.Picture
               sServPart(Grd.row, 3) = "0"
            End If
            sServPart(Grd.row, 4) = Format$(!OPSVCUNIT, ES_QuantityDataFormat)
            .MoveNext
         Loop
         ClearResultSet RdoRte
      End With
   End If
   Set RdoRte = Nothing
   
End Sub


Private Function GetMoQty() As Currency
   Dim RdoQty As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT RUNREMAININGQTY FROM RunsTable WHERE " _
          & "RUNREF='" & Compress(lblMon) & "' AND RUNNO=" _
          & Val(lblRun) & " "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoQty, ES_FORWARD)
   If bSqlRows Then
      With RdoQty
         If Not IsNull(!RUNREMAININGQTY) Then
            GetMoQty = !RUNREMAININGQTY
         Else
            GetMoQty = 0
         End If
         ClearResultSet RdoQty
      End With
   Else
      GetMoQty = 0
   End If
   Set RdoQty = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getmoqty"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub grd_KeyPress(KeyAscii As Integer)
   Dim bResponse As Byte
   Dim iPAonDock As Integer
   On Error Resume Next
   If sServPart(Grd.row, 0) = "" Then Exit Sub
   If KeyAscii = vbKeyReturn Or KeyAscii = vbKeySpace Then
      If sServPart(Grd.row, 3) = 1 Then
         MsgBox "That Item Has Been Purchased.", _
            vbInformation, Caption
      Else
      
         If (GetCompanyOPServSetting = True) Then
            Dim strMO As String
            Dim run As Integer
            Dim opNo As Integer
            Dim bRet As Boolean
            
            strMO = PurcPRe02c.cmbMon
            run = PurcPRe02c.cmbRun
            opNo = sServPart(Grd.row, 0)
            bRet = FindAnyOpenServiceParts(strMO, run, opNo)
            
            If (bRet = True) Then
               bResponse = MsgBox("There is an open Service Part operation remaining?", vbCritical, Caption)
               Exit Sub
            End If
            
         End If
         
         bResponse = MsgBox("Mark This Item As Purchased?", ES_YESQUESTION, Caption)
         If bResponse = vbYes Then
            PurcPRe02c.lblOpno = sServPart(Grd.row, 0)
            PurcPRe02c.lblServPart = sServPart(Grd.row, 1)
            PurcPRe02c.lblDsc = sServPart(Grd.row, 2)
            PurcPRe02c.txtPrc = sServPart(Grd.row, 4)
            PurcPRe02c.txtQty = Format(Val(lblRem), ES_QuantityDataFormat)
            PurcPRe02c.Grd.Col = 2
            PurcPRe02c.Grd.Text = PurcPRe02c.lblServPart
            PurcPRe02c.Grd.Col = 3
            PurcPRe02c.Grd.Text = PurcPRe02c.txtQty
            PurcPRe02c.cmbMon.Enabled = False
            PurcPRe02c.cmbRun.Enabled = False
            iPAonDock = GetPAOnDock(sServPart(Grd.row, 1))
            PurcPRe02c.optIns.Value = iPAonDock
            bItemSelected = 1
            UpdateItem
            Unload Me
         End If
      End If
   End If
   
End Sub

Private Function FindAnyOpenServiceParts(strMO As String, run As Integer, opNo As Integer) As Boolean
   Dim RdoOpSer As ADODB.Recordset
   On Error GoTo DiaErr1
   FindAnyOpenServiceParts = False
   
   sSql = "select OPSERVPART, OPNO from rnopTable where opref = '" & Compress(strMO) & "' and oprun = " & run _
            & " and opcomplete <> 1 and OPNO < " & opNo & " and OPSERVPART <> ''"
   
          
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoOpSer, ES_FORWARD)
   If bSqlRows Then
      With RdoOpSer
         If (!OPSERVPART <> "") Then
            FindAnyOpenServiceParts = True
         End If
         ClearResultSet RdoOpSer
      End With
   End If
   Set RdoOpSer = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "FindAnyOpenServiceParts"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Function GetCompanyOPServSetting() As Boolean
   Dim RdoOpSer As ADODB.Recordset
   On Error GoTo DiaErr1
   GetCompanyOPServSetting = False
   sSql = "SELECT ISNULL(COWARNSERVICEOPOPEN, 0)COWARNSERVICEOPOPEN FROM ComnTable"
          
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoOpSer, ES_FORWARD)
   If bSqlRows Then
      With RdoOpSer
         If (!COWARNSERVICEOPOPEN = 1) Then
            GetCompanyOPServSetting = True
         End If
         ClearResultSet RdoOpSer
      End With
   End If
   Set RdoOpSer = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "GetCompanyOPServSetting"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function


Private Function GetPAOnDock(sServPart As String) As Integer
   Dim RdoPADock As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT PAONDOCK FROM PartTable WHERE " _
          & "PartRef='" & Compress(sServPart) & "'"
          
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPADock, ES_FORWARD)
   If bSqlRows Then
      With RdoPADock
         If Not IsNull(!PAONDOCK) Then
            GetPAOnDock = !PAONDOCK
         Else
            GetPAOnDock = 0
         End If
         ClearResultSet RdoPADock
      End With
   Else
      GetPAOnDock = 0
   End If
   Set RdoPADock = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getmoqty"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function


Private Sub UpdateItem()
   On Error Resume Next
   clsADOCon.BeginTrans
   clsADOCon.ADOErrNum = 0
   
   sSql = "UPDATE PoitTable SET PIPART='" & Compress(PurcPRe02c.lblServPart) & "'," _
          & "PIRUNPART='" & Compress(PurcPRe02c.cmbMon) & "'," _
          & "PIRUNNO=" & Val(PurcPRe02c.cmbRun) & "," _
          & "PIRUNOPNO=" & Val(PurcPRe02c.lblOpno) & "," _
          & "PIPQTY=" & Val(PurcPRe02c.txtQty) & "," _
          & "PIONDOCK=" & Val(PurcPRe02c.optIns.Value) & "," _
          & "PIESTUNIT=" & Val(PurcPRe02c.txtPrc) & "," _
          & "PIPDATE='" & PurcPRe02c.txtDue & "', " _
          & "PIPORIGDATE='" & PurcPRe02c.txtDue & "' " _
          & "WHERE (PINUMBER=" & Val(PurcPRe02c.lblPon) & " AND " _
          & "PIITEM=" & Val(PurcPRe02c.lblItm) & ")"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "UPDATE RnopTable SET OPPONUMBER=" & Val(PurcPRe02c.lblPon) & "," _
          & "OPPOITEM=" & Val(PurcPRe02c.lblItm) & "," _
          & "OPPURCHASED=1 WHERE (OPREF='" & Compress(PurcPRe02c.cmbMon) & "' " _
          & "AND OPRUN=" & Val(PurcPRe02c.cmbRun) & " " _
          & "AND OPNO=" & Val(PurcPRe02c.lblOpno) & ")"
   clsADOCon.ExecuteSQL sSql
   If clsADOCon.ADOErrNum = 0 Then
      clsADOCon.CommitTrans
      PurcPRe02c.GetThisItem
   Else
      clsADOCon.RollbackTrans
   End If
End Sub

Private Sub Grd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim bResponse As Byte
   Dim iPAonDock As Integer
   On Error Resume Next
   If sServPart(Grd.row, 0) = "" Then Exit Sub
   If sServPart(Grd.row, 3) = 1 Then
      MsgBox "That Item Has Been Purchased.", _
         vbInformation, Caption
   Else
   
      If (GetCompanyOPServSetting = True) Then
         Dim strMO As String
         Dim run As Integer
         Dim opNo As Integer
         Dim bRet As Boolean
         
         strMO = PurcPRe02c.cmbMon
         run = PurcPRe02c.cmbRun
         opNo = sServPart(Grd.row, 0)
         bRet = FindAnyOpenServiceParts(strMO, run, opNo)
         
         If (bRet = True) Then
            bResponse = MsgBox("There is an open Service Part operation remaining?", vbCritical, Caption)
            Exit Sub
         End If
         
      End If

      bResponse = MsgBox("Mark This Item As Purchased?", ES_YESQUESTION, Caption)
      If bResponse = vbYes Then
         PurcPRe02c.lblOpno = sServPart(Grd.row, 0)
         PurcPRe02c.lblServPart = sServPart(Grd.row, 1)
         PurcPRe02c.lblDsc = sServPart(Grd.row, 2)
         PurcPRe02c.txtPrc = sServPart(Grd.row, 4)
         PurcPRe02c.txtQty = Format(Val(lblRem), ES_QuantityDataFormat)
         PurcPRe02c.Grd.Col = 2
         PurcPRe02c.Grd.Text = PurcPRe02c.lblServPart
         PurcPRe02c.Grd.Col = 3
         PurcPRe02c.Grd.Text = PurcPRe02c.txtQty
         PurcPRe02c.cmbMon.Enabled = False
         PurcPRe02c.cmbRun.Enabled = False
         iPAonDock = GetPAOnDock(sServPart(Grd.row, 1))
         PurcPRe02c.optIns.Value = iPAonDock
         bItemSelected = 1
         UpdateItem
         Unload Me
      End If
   End If
   
End Sub
