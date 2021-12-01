VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Begin VB.Form PickMCe03a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Scrap/Restock Information"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7455
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   7455
   ShowInTaskbar   =   0   'False
   Begin VB.Frame z2 
      Height          =   30
      Left            =   120
      TabIndex        =   25
      Top             =   1560
      Width           =   7212
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "PickMCe03a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   24
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CheckBox optLots 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   255
      Left            =   3120
      TabIndex        =   22
      Top             =   2760
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmdStk 
      Caption         =   "&Restock"
      Height          =   315
      Left            =   6480
      TabIndex        =   4
      ToolTipText     =   "Return The Quantity To Inventory"
      Top             =   2040
      Width           =   915
   End
   Begin VB.CommandButton cmdScr 
      Caption         =   "&Scrap"
      Height          =   315
      Left            =   6480
      TabIndex        =   5
      ToolTipText     =   "Mark The Quantity As Scrap"
      Top             =   2400
      Width           =   915
   End
   Begin VB.ComboBox cmbPpr 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1080
      TabIndex        =   2
      ToolTipText     =   "Picked Part Number"
      Top             =   2040
      Width           =   3255
   End
   Begin VB.TextBox txtQty 
      Height          =   285
      Left            =   4440
      TabIndex        =   3
      ToolTipText     =   "Adjustment Quantity"
      Top             =   2400
      Width           =   1215
   End
   Begin VB.ComboBox cmbRun 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   5400
      TabIndex        =   1
      ToolTipText     =   "Select Run Number"
      Top             =   720
      Width           =   1095
   End
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   1080
      TabIndex        =   0
      ToolTipText     =   "Select Part Number"
      Top             =   720
      Width           =   3255
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   6480
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   915
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6840
      Top             =   3000
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3330
      FormDesignWidth =   7455
   End
   Begin VB.Label Z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Lot Tracked"
      Height          =   255
      Index           =   7
      Left            =   1800
      TabIndex        =   23
      Top             =   2760
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Type "
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   21
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label lblTyp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1080
      TabIndex        =   20
      Top             =   2760
      Width           =   375
   End
   Begin VB.Label lblQty 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   4440
      TabIndex        =   19
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label lblRev 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   120
      TabIndex        =   18
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label lblUom 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5760
      TabIndex        =   17
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label lblPsc 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1080
      TabIndex        =   16
      Top             =   2400
      Width           =   3015
   End
   Begin VB.Label Z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Rev    "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   15
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Picked Part Number                                                 "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   1080
      TabIndex        =   14
      Top             =   1800
      Width           =   3015
   End
   Begin VB.Label Z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Pick/Adjust Qty           "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   4440
      TabIndex        =   13
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Uom     "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   5760
      TabIndex        =   12
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Status (PP,PC,CO)"
      Height          =   255
      Index           =   15
      Left            =   4440
      TabIndex        =   11
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1080
      TabIndex        =   10
      Top             =   1080
      Width           =   3255
   End
   Begin VB.Label lblStat 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   6000
      TabIndex        =   9
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Run"
      Height          =   255
      Index           =   1
      Left            =   4440
      TabIndex        =   8
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   720
      Width           =   1095
   End
End
Attribute VB_Name = "PickMCe03a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Stanwood, Washington, USA  ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***

Option Explicit
Dim AdoQry As ADODB.Command
Dim AdoQParameter1 As ADODB.Parameter

Dim AdoItm As ADODB.Command
Dim AdoIParameter1 As ADODB.Parameter
Dim AdoIParameter2 As ADODB.Parameter

Dim bCancel As Byte
Dim bOnLoad As Byte
Dim bGoodPick As Byte
Dim bGoodRuns As Byte

Dim iIndex As Integer
Dim iOldIndex As Integer
Dim iOldRun As Integer
Dim iTotalItems As Integer

Dim sPartNumber As String
Dim sCreditAcct As String
Dim sDebitAcct As String
Dim cUnitCost As Currency

Dim vItems(2000, 11) As Variant
Private Const PICK_UnitCost = 3
Private Const PICK_LOTTRACKED = 8
Private Const PICK_USEACTUALCOST = 10
Private Const PICK_RecNo = 9

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub ScrapThese()
   Dim bResponse As Byte
   Dim iPkRecord As Integer
   
   Dim sMsg As String
   Dim sDate As String
   
   On Error Resume Next
   sDate = Format(ES_SYSDATE, "mm/dd/yy")
   If Val(txtQty) = 0 Then
      MsgBox "You Have Entered a Zero Quantity.", 64, Caption
      txtQty.SetFocus
      Exit Sub
   Else
      sMsg = "You Have Chosen To Scrap " & txtQty & " " & lblUom & vbCr _
             & "Part Number " & vItems(iIndex, 1) & "." & vbCr _
             & "Do You Wish To Continue?"
      bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
      If bResponse = vbYes Then
         cmbPpr.Enabled = False
         txtQty.Enabled = False
         cmdScr.Enabled = False
         cmdStk.Enabled = False
         MouseCursor 13
         clsADOCon.BeginTrans
         iPkRecord = GetNextPickRecord(sPartNumber, Val(cmbRun))
         sSql = "INSERT INTO MopkTable (PKPARTREF,PKMOPART,PKMORUN,PKTYPE," _
                & "PKREV,PKPDATE,PKADATE,PKPQTY,PKAQTY,PKRECORD,PKUNITS,PKCOMT) " _
                & "VALUES('" & vItems(iIndex, 0) & "','" & sPartNumber & "'," _
                & cmbRun & "," & IATYPE_PickScrap & ",'" & lblRev & "','" & sDate & "','" & sDate & "'," _
                & txtQty & "," & txtQty & "," & iPkRecord & ",'" _
                & lblUom & "','Reduction/Scrap')"
         clsADOCon.ExecuteSql sSql
         
         sSql = "UPDATE MopkTable SET PKAQTY=PKAQTY-" & txtQty & " " _
                & "WHERE PKPARTREF='" & vItems(iIndex, 0) & "' AND " _
                & "PKMOPART='" & sPartNumber & "' AND " _
                & "PKMORUN=" & cmbRun & " AND " _
                & "PKRECORD=" & vItems(iIndex, PICK_RecNo) & "'"
         clsADOCon.ExecuteSql sSql
         clsADOCon.CommitTrans
         MouseCursor 0
         MsgBox "Quantity Was Successfully Scrapped.", 64, Caption
         On Error Resume Next
         cmbRun.SetFocus
      Else
         CancelTrans
      End If
   End If
   Exit Sub
   
DiaErr1:
   MouseCursor 0
   On Error Resume Next
   clsADOCon.RollbackTrans
   sMsg = CurrError.Description & vbCr _
          & "Could Not Complete Pick Scrap."
   MsgBox sMsg, 48, Caption
   
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtQty = "0.000"
   b = CheckLotStatus()
   If b = 1 Then
      z1(7).Visible = True
      optLots.Visible = True
   End If
   
End Sub

Private Sub cmbPpr_Click()
   On Error Resume Next
   iIndex = cmbPpr.ListIndex
   If iIndex < 0 Then iIndex = 0
   lblPsc = vItems(iIndex, 2)
   lblQty = vItems(iIndex, 6)
   lblUom = vItems(iIndex, 4)
   lblTyp = vItems(iIndex, 7)
   optLots.Value = vItems(iIndex, PICK_LOTTRACKED)
   txtQty = ""
   iOldIndex = iIndex
End Sub

Private Sub cmbPpr_GotFocus()
   cmbPpr_Click
End Sub

Private Sub cmbPpr_LostFocus()
   cmbPpr = CheckLen(cmbPpr, 30)
   On Error GoTo MrstkCp1
   iIndex = cmbPpr.ListIndex
   cmbPpr = vItems(iIndex, 1)
   lblPsc = vItems(iIndex, 2)
   lblQty = vItems(iIndex, 6)
   lblUom = vItems(iIndex, 4)
   lblTyp = vItems(iIndex, 7)
   optLots.Value = vItems(iIndex, PICK_LOTTRACKED)
   Exit Sub
   
MrstkCp1:
   Resume MrstkCp2
MrstkCp2:
   'Beep
   On Error Resume Next
   If iOldIndex < 0 Then iOldIndex = 0
   iIndex = iOldIndex
   cmbPpr = vItems(iIndex, 1)
   lblPsc = vItems(iIndex, 2)
   lblQty = vItems(iIndex, 6)
   lblUom = vItems(iIndex, 4)
   
End Sub


Private Sub cmbPrt_Click()
   optLots.Value = vbUnchecked
   bGoodRuns = GetRuns()
   
End Sub


Private Sub cmbPrt_GotFocus()
   cmbPrt_Click
   
End Sub


Private Sub cmbPrt_LostFocus()
   If bCancel Then Exit Sub
   
   If (Not ValidPartNumber(cmbPrt.Text)) Then
      MsgBox "Can't Select The Part Number Which Is Obsolete or Inactive. ", _
         vbInformation, Caption
      cmbPrt = ""
      Exit Sub
   End If
   
   cmbPrt = CheckLen(cmbPrt, 30)
   
   bGoodRuns = GetRuns()
   
End Sub


Private Sub cmbRun_Click()
   bGoodPick = GetPick()
   
End Sub


Private Sub cmbRun_LostFocus()
   If bCancel Then Exit Sub
   cmbRun = CheckLen(cmbRun, 5)
   cmbRun = Format(Abs(Val(cmbRun)), "####0")
   bGoodPick = GetPick()
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub



Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext "5206"
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub cmdScr_Click()
   ScrapThese
   
End Sub

Private Sub cmdStk_Click()
   RestockThese
   
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillCombo
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   
   sSql = "SELECT PARTREF,PARTNUM,PADESC,PALEVEL,PASTDCOST,PAUNITS," _
          & "PALOTTRACK,PKPARTREF,PKMOPART,PKMORUN,PKREV,PKTYPE,PKAQTY," _
          & "PKRECORD,PAUSEACTUALCOST FROM PartTable,MopkTable WHERE PARTREF=PKPARTREF " _
          & "AND PKMOPART= ? AND PKMORUN= ? " _
          & "AND PKAQTY>0 AND PKTYPE=10 ORDER BY PARTREF"
   'Set RdoItm = RdoCon.CreateQuery("", sSql)
   Set AdoItm = New ADODB.Command
   AdoItm.CommandText = sSql
   Set AdoIParameter1 = New ADODB.Parameter
   AdoIParameter1.Size = 30
   AdoIParameter1.Type = adChar
   AdoItm.Parameters.Append AdoIParameter1
   
   Set AdoIParameter2 = New ADODB.Parameter
   AdoIParameter2.Type = adInteger
   AdoItm.Parameters.Append AdoIParameter2
   
   
   sSql = "SELECT RUNREF,RUNSTATUS,RUNNO FROM " _
          & "RunsTable WHERE RUNREF = ? " _
          & "AND (RUNSTATUS='PP' OR RUNSTATUS='PC' OR RUNSTATUS='CO')"
   'Set RdoQry = RdoCon.CreateQuery("", sSql)
   Set AdoQry = New ADODB.Command
   AdoQry.CommandText = sSql
   
   Set AdoQParameter1 = New ADODB.Parameter
   AdoQParameter1.Type = adChar
   AdoQParameter1.Size = 30
   AdoQry.Parameters.Append AdoQParameter1
   
   
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set AdoIParameter1 = Nothing
   Set AdoIParameter2 = Nothing
   Set AdoItm = Nothing
   Set AdoQParameter1 = Nothing
   Set AdoQry = Nothing
   Set PickMCe03a = Nothing
   
End Sub



Private Sub FillCombo()
   Dim RdoCmb As ADODB.Recordset
   Dim b As Byte
   Dim sTempPart As String
   
   On Error GoTo DiaErr1
   sJournalID = GetOpenJournal("IJ", Format$(ES_SYSDATE, "mm/dd/yy"))
   If Left(sJournalID, 4) = "None" Then
      sJournalID = ""
      b = 1
   Else
      If sJournalID = "" Then b = 0 Else b = 1
   End If
   If b = 0 Then
      MsgBox "There Is No Open Inventory Journal For This Period.", _
         vbExclamation, Caption
      Sleep 500
      Unload Me
      Exit Sub
   End If
   sProcName = "fillcombo"
   sSql = "SELECT DISTINCT PARTREF,PARTNUM,PADESC,RUNREF " _
          & "PKMOPART,PKAQTY FROM PartTable,RunsTable,MopkTable " _
          & "MopkTable WHERE PARTREF=RUNREF AND (PARTREF=PKMOPART AND PKMORUN=RUNNO)" _
          & "AND (RUNSTATUS='PP' OR RUNSTATUS='PC' OR RUNSTATUS='CO') AND PKAQTY>0 " _
          & "ORDER BY PARTREF"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCmb, ES_FORWARD)
   If bSqlRows Then
      With RdoCmb
         cmbPrt = "" & Trim(!PartNum)
         lblDsc = "" & Trim(!PADESC)
         Do Until .EOF
            If sTempPart <> Trim(!PartNum) Then
               AddComboStr cmbPrt.hWnd, "" & Trim(!PartNum)
               sTempPart = Trim(!PartNum)
            End If
            .MoveNext
         Loop
         ClearResultSet RdoCmb
      End With
   End If
   Set RdoCmb = Nothing
   bGoodRuns = GetRuns()
   Exit Sub
   
DiaErr1:
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Function GetRuns() As Byte
   Dim RdoRun As ADODB.Recordset
   If bCancel = 1 Then Exit Function
   If sPartNumber = Compress(cmbPrt) Then
      Exit Function
   End If
   cmbRun.Clear
   ClearBoxes
   sPartNumber = Compress(cmbPrt)
   FindPart cmbPrt
   lblTyp = ""
   On Error GoTo DiaErr1
   AdoQry.Parameters(0).Value = sPartNumber
   bSqlRows = clsADOCon.GetQuerySet(RdoRun, AdoQry)
   If bSqlRows Then
      With RdoRun
         cmbRun = Format(!Runno, "####0")
         lblStat = "" & !RUNSTATUS
         Do Until .EOF
            AddComboStr cmbRun.hWnd, Format$(!Runno, "####0")
            .MoveNext
         Loop
         ClearResultSet RdoRun
      End With
      GetRuns = 1
   Else
      sPartNumber = ""
      GetRuns = 0
   End If
   Set RdoRun = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getruns"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Function GetPick() As Byte
   MouseCursor 13
   Dim RdoPck As ADODB.Recordset
   Dim iList As Integer
   
   Erase vItems
   ClearBoxes
   iOldRun = Val(cmbRun)
   iList = -1
   iIndex = 0
   On Error GoTo DiaErr1

   AdoItm.Parameters(0).Value = sPartNumber
   AdoItm.Parameters(1).Value = Val(cmbRun)
   
   bSqlRows = clsADOCon.GetQuerySet(RdoPck, AdoItm, ES_KEYSET, True)
   If bSqlRows Then
      With RdoPck
         'set controls to values for first pick item
         cmbPpr = "" & Trim(!PartNum)
         lblTyp = Format(0 + !PALEVEL, "0")
         lblPsc = "" & Trim(!PADESC)
         lblRev = "" & Trim(!PKREV)
         lblQty = Format(!PKAQTY, ES_QuantityDataFormat)
         lblUom = "" & Trim(!PAUNITS)
         cmbPpr.Enabled = True
         txtQty.Enabled = True
         cmdStk.Enabled = True
         cmdScr.Enabled = True
         If optLots.Visible Then
            optLots.Value = !PALOTTRACK
         Else
            optLots.Value = vbUnchecked
         End If
         
         Do Until .EOF
            iList = iList + 1
            AddComboStr cmbPpr.hWnd, "" & Trim(!PartNum)
            vItems(iList, 0) = "" & Trim(!PartRef)
            vItems(iList, 1) = "" & Trim(!PartNum)
            vItems(iList, 2) = "" & Trim(!PADESC)
            vItems(iList, PICK_UnitCost) = Format(!PASTDCOST, ES_QuantityDataFormat)
            vItems(iList, PICK_USEACTUALCOST) = "" & Trim(!PAUSEACTUALCOST)
            vItems(iList, 4) = "" & Trim(!PAUNITS)
            vItems(iList, 5) = "" & Trim(!PKREV)
            vItems(iList, 6) = Format(!PKAQTY, ES_QuantityDataFormat)
            vItems(iList, 7) = Format(!PALEVEL, "0")
            If optLots.Visible Then
               vItems(iList, PICK_LOTTRACKED) = Format(!PALOTTRACK, "0")
            Else
               vItems(iList, PICK_LOTTRACKED) = "0"
            End If
            
            '.Edit
            'If !PKRECORD = 0 Then !PKRECORD = iList + 1
            ' TODO: Edit is not there so the recordset is not
            ' Opened for editing (You have to use AddNew to change to editmode
            .Update
            
            vItems(iList, PICK_RecNo) = Format(!PKRECORD, "##0")
            .MoveNext
         Loop
         ClearResultSet RdoPck
      End With
      GetPick = True
      iTotalItems = iList - 1
   Else
      iTotalItems = 0
      GetPick = False
      ClearBoxes
      MouseCursor 0
      MsgBox "There Are No Picked Items.", vbInformation, Caption
   End If
   On Error Resume Next
   Set RdoPck = Nothing
   MouseCursor 0
   Exit Function
   
DiaErr1:
   sProcName = "getpick"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub ClearBoxes()
   cmbPpr.Clear
   lblRev = ""
   lblPsc = ""
   lblQty = ""
   lblUom = ""
   lblTyp = ""
   txtQty = ""
   optLots.Value = vbUnchecked
   cmdStk.Enabled = False
   cmdScr.Enabled = False
   
End Sub

Private Sub lblDsc_Change()
   If Left(lblDsc, 8) = "*** Part" Then
      lblDsc.ForeColor = ES_RED
   Else
      lblDsc.ForeColor = vbBlack
   End If
   
End Sub

Private Sub txtQty_LostFocus()
   If Val(txtQty) > Val(lblQty) Then
      'Beep
      txtQty = lblQty
   End If
   txtQty = CheckLen(txtQty, PICK_RecNo)
   txtQty = Format(Abs(Val(txtQty)), ES_QuantityDataFormat)
   
End Sub




Public Sub RestockThese()
   Dim bResponse As Byte
   Dim bLotsFail As Byte
   
   Dim A As Integer
   Dim iList As Integer
   Dim iPkRecord As Integer
   
   Dim lCOUNTER As Long
   Dim lSysCount As Long
   Dim lLOTRECORD As Long
   Dim lRun As Long
   
   Dim cCost As Currency
   Dim cLotQty As Currency
   Dim cQuantity As Currency
   Dim cReStockMatl As Currency

   Dim sMsg As String
   Dim sLotNumber As String
   Dim sMon As String
   Dim sMoRun As String * 9
   Dim sMoPart As String * 31
   Dim sNewPart As String
   
   Dim sDate As Variant
   'On Error Resume Next
   sDate = Format(ES_SYSDATE, "mm/dd/yy hh:mm")
   cQuantity = Format("0" & txtQty, "#########0.000")
   If Val(txtQty) > Val(lblQty) Then
      MsgBox "You Have Entered a Quantity Greater Than Picked.", 64, Caption
      txtQty.SetFocus
   End If
   If Val(txtQty) = 0 Then
      MsgBox "You Have Entered a Zero Quantity.", 64, Caption
      txtQty.SetFocus
      Exit Sub
   End If
      
   sMsg = "You Have Chosen To Restock " & txtQty _
          & " For " & vItems(iIndex, 1) & "" & vbCr _
          & "To Inventory. Do You Wish To Continue?"
   bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
   If bResponse <> vbYes Then
      Exit Sub
   End If
      
   'determine how to allocate restocks to lots
   sNewPart = vItems(iIndex, 0)
   sMon = Compress(cmbPrt)                   'use for SQL
   lRun = Val(cmbRun)                        'use for SQL
   sMoPart = cmbPrt                          'formatted for comments
   sMoRun = "RUN" & Space$(iList) & cmbRun   'formatted for comments
   MouseCursor 13
   cmbPpr.Enabled = False
   txtQty.Enabled = False
   cmdStk.Enabled = False
   cmdScr.Enabled = False
   
   'if lot tracked, allow user to select lots to restock to
   If optLots.Value = vbChecked Then
      PickMCe03b.lblMon = cmbPrt
      PickMCe03b.lblRun = lRun
      PickMCe03b.lblPart = cmbPpr
      PickMCe03b.lblRestockQty = Format(cQuantity, ES_QuantityDataFormat) 'cQuantity
      PickMCe03b.Show vbModal
      If Es_LotSelectionCanceled Then
         'RdoCon.RollbackTrans
         Exit Sub
      End If
            
   'if no lot tracking, put the quantities back randomly to the lots from
   'which the parts were drawn.  (There should still be lots underneath, even if the
   'part is not lot tracked)
   Else
      Dim lot As New ClassLot
      lot.AllocatePickRestocks cmbPrt, lRun, cmbPpr, cQuantity
   End If
         
   'begin transaction
   clsADOCon.BeginTrans
   
   lCOUNTER = GetLastActivity()
   lSysCount = lCOUNTER + 1
   iList = Len(Trim(str(cmbRun)))
   iList = 5 - iList
   iPkRecord = GetNextPickRecord(sMon, lRun)
   
   bResponse = GetPartAccounts(sNewPart, sDebitAcct, sCreditAcct)
   sSql = "UPDATE MopkTable SET PKAQTY=PKAQTY-" & cQuantity & " " _
          & "WHERE PKPARTREF='" & sNewPart & "' AND " _
          & "PKMOPART='" & sMon & "' AND " _
          & "PKMORUN=" & lRun & " AND " _
          & "PKRECORD=" & vItems(iIndex, PICK_RecNo) & " "
   clsADOCon.ExecuteSql sSql
   
   'return quantities to lots
   If Es_TotalLots > 0 Then
      For A = 1 To UBound(lots)
        
        ' If ActualCost get the unit cost from Lot header table,
        ' Else use the parttable cost value.
        If (vItems(iIndex, PICK_USEACTUALCOST) = "1") Then
            Dim lot1 As New ClassLot
            cUnitCost = lot1.GetLotUnitCost(sNewPart, lots(A).LotSysId)
        Else
            cUnitCost = vItems(iIndex, PICK_UnitCost)
        End If

         'insert lot transaction here
         lCOUNTER = lCOUNTER + 1
         lLOTRECORD = GetNextLotRecord(lots(A).LotSysId)
         sSql = "INSERT INTO LoitTable (LOINUMBER,LOIRECORD," _
                & "LOITYPE,LOIPARTREF,LOIQUANTITY," _
                & "LOIMOPARTREF,LOIMORUNNO," _
                & "LOIACTIVITY,LOICOMMENT) " _
                & "VALUES('" & lots(A).LotSysId & "'," _
                & lLOTRECORD & "," & IATYPE_PickRestock & ",'" & sNewPart & "'," _
                & lots(A).LotSelQty & ",'" & sMon & "'," & lRun & "," _
                & lCOUNTER & ",'Pick Return To Stock')"
         clsADOCon.ExecuteSql sSql
         
         ' Calculate the Total Material cost.
         cReStockMatl = cUnitCost * lots(A).LotSelQty

         ' 7/9/2016 - Swapped the Credit/Debit columns
         sSql = "INSERT INTO InvaTable" & vbCrLf _
            & "(INTYPE,INPART,INREF1,INREF2,INPQTY,INAQTY,INPDATE,INADATE," & vbCrLf _
            & "INAMT,INTOTMATL,INDEBITACCT,INCREDITACCT,INMOPART,INMORUN,INNUMBER,INLOTNUMBER,INUSER) " & vbCrLf _
            & "VALUES(" & IATYPE_PickRestock & ",'" & sNewPart & "'," _
            & "'RETURN TO STOCK','" & sMoPart & sMoRun _
            & "'," & lots(A).LotSelQty & "," & lots(A).LotSelQty & "," & vbCrLf _
            & "'" & sDate & "','" & sDate & "'," & Format(cUnitCost, ES_MoneyFormat) & "," _
            & "-" & Format(cReStockMatl, ES_MoneyFormat) & ",'" & sCreditAcct & "','" & sDebitAcct & "','" _
            & sMon & "'," & lRun & "," & lCOUNTER & ",'" & lots(A).LotSysId & "','" & sInitials & "')"
         clsADOCon.ExecuteSql sSql
         
         sSql = "UPDATE LohdTable SET LOTREMAININGQTY=LOTREMAININGQTY" _
                & "+" & lots(A).LotSelQty & " WHERE LOTNUMBER='" _
                & lots(A).LotSysId & "'"
         clsADOCon.ExecuteSql sSql
      Next
   End If
         
   'commented out -- there should always be lots
''   'if no lots, create one
''   If optLots.Value = vbUnchecked Then
''      lCOUNTER = lCOUNTER + 1
''      sSql = "INSERT INTO InvaTable" & vbCrLf _
''         & "(INTYPE,INPART,INREF1,INREF2,INPQTY,INAQTY,INPDATE,INADATE," & vbCrLf _
''         & "INAMT,INCREDITACCT,INDEBITACCT,INMOPART,INMORUN,INNUMBER,INUSER,INUNITS)" & vbCrLf _
''         & "VALUES(" & IATYPE_PickRestock & ",'" & sNewPart & "'," _
''         & "'RETURN TO STOCK','" & sMoPart & sMoRun _
''         & "'," & cQuantity & "," & cQuantity & ",'" & sDate & "','" & sDate & "'," & vbCrLf _
''         & vItems(iIndex, PICK_UnitCost) & ",'" _
''         & sCreditAcct & "','" & sDebitAcct & "','" & sMon & "'," & lRun & "," _
''         & lCOUNTER & ",'" & sInitials & "','" & lblUom & "')"
''      RdoCon.Execute sSql, rdExecDirect
''
''      'New lot
''      cLotQty = Val(txtQty)
''      sLotNumber = GetNextLotNumber()
''      sSql = "INSERT INTO LohdTable (LOTNUMBER,LOTUSERLOTID," _
''             & "LOTPARTREF,LOTPDATE,LOTORIGINALQTY,LOTREMAININGQTY," _
''             & "LOTUNITCOST,LOTDATECOSTED) " _
''             & "VALUES('" _
''             & sLotNumber & "','Return To Stock-" & sLotNumber & "','" & sMon _
''             & "','" & sDate & "'," & Trim(Str(cLotQty)) & "," & Trim(Str(cLotQty)) _
''             & "," & cCost & ",'" & sDate & "')"
''      RdoCon.Execute sSql, rdExecDirect
''
''      sSql = "INSERT INTO LoitTable (LOINUMBER,LOIRECORD," _
''             & "LOITYPE,LOIPARTREF,LOIPDATE,LOIQUANTITY," _
''             & "LOIACTIVITY,LOICOMMENT,LOIUNITS) " _
''             & "VALUES('" _
''             & sLotNumber & "',1," & IATYPE_PickRestock & ",'" & sMon _
''             & "','" & sDate & "'," & Trim(Str(cLotQty)) _
''             & "," & lCOUNTER & ",'" _
''             & "Return To Stock" & "','" & lblUom & "')"
''      RdoCon.Execute sSql, rdExecDirect
''   End If


   sSql = "UPDATE PartTable SET PAQOH=PAQOH+" & cQuantity & "," _
          & "PALOTQTYREMAINING=PALOTQTYREMAINING+" & cQuantity & " " _
          & "WHERE PARTREF='" & sNewPart & "' "
   clsADOCon.ExecuteSql sSql
   AverageCost sNewPart
   UpdateWipColumns lSysCount
   clsADOCon.CommitTrans
   MouseCursor 0
   MsgBox "Material Restock Completed Successfully.", 64, Caption
   bGoodPick = GetPick()
   cmbRun.SetFocus
'   Else
'            RdoCon.RollbackTrans
'            MsgBox "The Material Restock Transaction Was Not Successful.", _
'               vbInformation, Caption
'            cmbRun.SetFocus
'         End If
'      Else
'         CancelTrans
'         txtQty = lblQty
'      End If
'   End If
   Exit Sub
   
DiaErr1:
'   On Error Resume Next
   clsADOCon.RollbackTrans
'   sMsg = CurrError.Description & vbCr _
'          & "Could Not Complete Pick Restock."
'   MsgBox sMsg, 48, Caption
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub
