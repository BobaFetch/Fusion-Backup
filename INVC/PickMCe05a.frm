VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form PickMCe05a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add A Pick List Item"
   ClientHeight    =   4920
   ClientLeft      =   1620
   ClientTop       =   960
   ClientWidth     =   6810
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   6810
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox optActualCost 
      Alignment       =   1  'Right Justify
      Caption         =   "Actual Cost"
      Enabled         =   0   'False
      Height          =   255
      Left            =   2400
      TabIndex        =   29
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Frame z2 
      ForeColor       =   &H8000000F&
      Height          =   40
      Left            =   60
      TabIndex        =   28
      Top             =   1660
      Width           =   6624
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "PickMCe05a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   27
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CheckBox optPick 
      Alignment       =   1  'Right Justify
      Caption         =   "Pick This Item"
      Height          =   255
      Left            =   4560
      TabIndex        =   26
      ToolTipText     =   "Pick Complete The New Item"
      Top             =   3360
      Value           =   1  'Checked
      Width           =   2055
   End
   Begin VB.TextBox txtCmt 
      Height          =   885
      Left            =   1200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      ToolTipText     =   "Optional Comments (2048 Char Max)"
      Top             =   3840
      Width           =   3855
   End
   Begin VB.CheckBox optPrompt 
      Alignment       =   1  'Right Justify
      Caption         =   "&Prompt To Add MO Or PO"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Prompts User To Switch To MO Or PO Entry (Workstation Setting)"
      Top             =   1400
      Value           =   1  'Checked
      Width           =   2295
   End
   Begin VB.CheckBox optLot 
      Alignment       =   1  'Right Justify
      Caption         =   "Lot Tracked Part"
      Enabled         =   0   'False
      Height          =   255
      Left            =   2640
      TabIndex        =   21
      Top             =   3000
      Width           =   1575
   End
   Begin VB.CommandButton cmdChg 
      Caption         =   "&Add"
      Height          =   315
      Left            =   5780
      TabIndex        =   6
      ToolTipText     =   "Add This Item To The Pick List (Or Create A Picklist)"
      Top             =   1740
      Width           =   915
   End
   Begin VB.TextBox txtQty 
      Height          =   285
      Left            =   4560
      TabIndex        =   4
      ToolTipText     =   "Pick Quantity"
      Top             =   2640
      Width           =   1095
   End
   Begin VB.ComboBox cmbPpr 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1200
      TabIndex        =   3
      ToolTipText     =   " Part Number To Be Picked (Qualifying Parts)"
      Top             =   2280
      Width           =   3255
   End
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   1200
      TabIndex        =   0
      ToolTipText     =   "Select Project Part Number"
      Top             =   720
      Width           =   3545
   End
   Begin VB.ComboBox cmbRun 
      ForeColor       =   &H00800000&
      Height          =   288
      Left            =   5640
      TabIndex        =   1
      ToolTipText     =   "Select Run Number"
      Top             =   720
      Width           =   1040
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5780
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   0
      Width           =   915
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   360
      Top             =   5160
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4920
      FormDesignWidth =   6810
   End
   Begin VB.Label Z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Pick Comments (Optional):"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   25
      Top             =   3480
      Width           =   2175
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   5640
      TabIndex        =   24
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      Height          =   255
      Index           =   2
      Left            =   4920
      TabIndex        =   23
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Add A Part To A Pick List Or Part Without A Pick List"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   360
      Width           =   4575
   End
   Begin VB.Label lblTyp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1680
      TabIndex        =   20
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Type "
      Height          =   255
      Index           =   7
      Left            =   1200
      TabIndex        =   19
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label lblCst 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1080
      TabIndex        =   18
      Top             =   5280
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Material"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   17
      Top             =   2280
      Width           =   1095
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
      Index           =   5
      Left            =   5760
      TabIndex        =   16
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Qoh/Chg Qty         "
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
      Left            =   4560
      TabIndex        =   15
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Z1 
      BackStyle       =   0  'Transparent
      Caption         =   "New Part Number                                                 "
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
      Left            =   1200
      TabIndex        =   14
      Top             =   2040
      Width           =   3015
   End
   Begin VB.Label lblPsc 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1200
      TabIndex        =   13
      Top             =   2640
      Width           =   3015
   End
   Begin VB.Label lblUom 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5760
      TabIndex        =   12
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label lblQty 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   4560
      TabIndex        =   11
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Run"
      Height          =   255
      Index           =   1
      Left            =   4920
      TabIndex        =   9
      Top             =   720
      Width           =   735
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1200
      TabIndex        =   8
      Top             =   1080
      Width           =   3255
   End
End
Attribute VB_Name = "PickMCe05a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'2/16/04 New
'3/4/04 Added Customer prompt for this copy only
'Add ability to add without picking
'9/1/04 omit tools
'6/28/05 Corrected Lots and Error in Inv activity (PickAddedPart)
'7/7/05 Corrected LoitTable.LOIMOPARTREF/LOIMORUN
'10/20/05 Corrected Distinct query (redundant Parts) FillCombo
'3/21/07 Corrected non lottracked part LoitTable inserts
Option Explicit
Dim AdoQry As ADODB.Command
Dim AdoParameter As ADODB.Parameter

Dim bFIFO As Byte
Dim bOnLoad As Byte
Dim bGoodMat As Byte
Dim bGoodRuns As Byte

Dim sPartNumber As String
Dim sCreditAcct As String
Dim sDebitAcct As String
Dim sJournalID As String

Dim sLots(50, 2) As String
'0 = Lot Number
'1 = Lot Quantity

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd


Private Sub GetThisRun()
   On Error GoTo DiaErr1
   Dim RdoStatus As ADODB.Recordset
   sSql = "SELECT RUNREF,RUNNO,RUNSTATUS FROM RunsTable " _
          & "WHERE (RUNREF='" & Compress(cmbPrt) & "' AND " _
          & "RUNNO=" & Val(cmbRun) & ")"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoStatus, ES_FORWARD)
   If bSqlRows Then lblStatus = "" & Trim(RdoStatus!RUNSTATUS) Else _
                                lblStatus = ""
   ClearResultSet RdoStatus
   Set RdoStatus = Nothing
   Exit Sub
DiaErr1:
   On Error GoTo 0
   
End Sub

Private Function GetPartLots(sPartWithLot As String) As Integer
   Dim RdoLots As ADODB.Recordset
   Dim iList As Integer
   Erase sLots
   On Error GoTo DiaErr1
   sSql = "SELECT LOTNUMBER,LOTPARTREF,LOTREMAININGQTY,LOTAVAILABLE " _
          & "FROM LohdTable WHERE (LOTPARTREF='" & sPartWithLot & "' AND " _
          & "LOTREMAININGQTY>0 AND LOTAVAILABLE=1) "
   If bFIFO = 1 Then
      sSql = sSql & "ORDER BY LOTNUMBER ASC"
   Else
      sSql = sSql & "ORDER BY LOTNUMBER DESC"
   End If
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoLots, ES_FORWARD)
   If bSqlRows Then
      With RdoLots
         Do Until .EOF
            iList = iList + 1
            sLots(iList, 0) = "" & Trim(!lotNumber)
            sLots(iList, 1) = Format$(!LOTREMAININGQTY, ES_QuantityDataFormat)
            .MoveNext
         Loop
         ClearResultSet RdoLots
      End With
      GetPartLots = iList
   Else
      GetPartLots = 0
   End If
   Set RdoLots = Nothing
   Exit Function
   
DiaErr1:
   GetPartLots = 0
   
End Function

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtQty = "0.000"
   If Left(MdiSect.Caption, 3) = "Inv" Then optPrompt.Visible = False
   
End Sub

Private Sub cmbPpr_Click()
   bGoodMat = FindMatPart()
   
End Sub

Private Sub cmbPpr_LostFocus()
   cmbPpr = CheckLen(cmbPpr, 30)
   
   If (Not ValidPartNumber(cmbPpr.Text)) Then
      MsgBox "Can't Select The Part Number Which Is Obsolete or Inactive. ", _
         vbInformation, Caption
      cmbPpr = ""
      Exit Sub
   End If
   
   bGoodMat = FindMatPart()
   
End Sub


Private Sub cmbPrt_Click()
   bGoodRuns = GetRuns()
   
End Sub


Private Sub cmbPrt_GotFocus()
   cmbPrt_Click
   
End Sub


Private Sub cmbPrt_LostFocus()
   cmbPrt = CheckLen(cmbPrt, 30)
   
   If (Not ValidPartNumber(cmbPrt.Text)) Then
      MsgBox "Can't Select The Part Number Which Is Obsolete or Inactive. ", _
         vbInformation, Caption
      cmbPrt = ""
      Exit Sub
   End If
   
   bGoodRuns = GetRuns()
   FillMaterial
   
End Sub


Private Sub cmbRun_Click()
   GetThisRun
   
End Sub


Private Sub cmbRun_LostFocus()
   cmbRun = Format(Abs(Val(cmbRun)), "####0")
   GetThisRun
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub



Private Sub cmdChg_Click()
   Dim bResponse As Byte
   If lblStatus = "CA" Or lblStatus = "CL" Or lblStatus = "CO" Or lblStatus = "" Then
      MsgBox "That Run Is Canceled, Complete, Or Closed.", _
         vbInformation, Caption
   Else
      If IsOpenPartAlreadyOnPickList(cmbPpr.Text) Then
         If MsgBox("Part " & cmbPpr.Text & " is already on the Pick List. " & vbCrLf & "Would you like to Revise the Quantity Now?", vbYesNo) = vbYes Then
            PickMCe01b.cbfrom1a = vbUnchecked
            PickMCe01b.lblMon = cmbPrt
            PickMCe01b.lblRun = cmbRun
'            PickMCe01a.cmbPrt = Me.cmbPrt
'            PickMCe01a.cmbRun = Me.cmbRun
            PickMCe01b.Show
            Unload Me
            Exit Sub
         End If
      End If
      
      If optPick.Value = vbChecked Then
         PickAddedPart
      Else
         AddPickItem
      End If
   End If
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 5205
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub


Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      bFIFO = GetInventoryMethod()
      FillMaterial
      FillCombo
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   GetOptions
'   sSql = "SELECT RUNREF,RUNSTATUS,RUNNO FROM " _
'          & "RunsTable WHERE RUNREF = ? " _
'          & "AND (RUNSTATUS<>'CA' OR RUNSTATUS<>'CL')"
   sSql = "SELECT RUNREF,RUNSTATUS,RUNNO FROM " _
          & "RunsTable WHERE RUNREF = ? " _
          & "AND RUNSTATUS NOT IN ('CA', 'CO','CL')"
   Set AdoQry = New ADODB.Command
   AdoQry.CommandText = sSql
   
   Set AdoParameter = New ADODB.Parameter
   AdoParameter.Size = 30
   AdoParameter.Type = adChar
   
   AdoQry.Parameters.Append AdoParameter
   
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
   Set AdoParameter = Nothing
   Set AdoQry = Nothing
   Set PickMCe05a = Nothing
   
End Sub



Private Sub FillCombo()
   Dim b As Byte
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
   
   sSql = "SELECT DISTINCT PARTREF,PARTNUM,RUNREF " & vbCrLf _
      & "FROM PartTable" & vbCrLf _
      & "JOIN RunsTable ON PARTREF=RUNREF" & vbCrLf _
      & "WHERE RUNSTATUS NOT IN ('CA', 'CO', 'CL')" & vbCrLf _
      & "ORDER BY PARTREF"
   LoadComboBox cmbPrt
   If bSqlRows Then
      cmbPrt = cmbPrt.List(0)
      bGoodRuns = GetRuns()
   Else
      MsgBox "No Runs Recorded.", vbInformation, Caption
   End If
   Exit Sub
   
DiaErr1:
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub



Private Function GetRuns() As Byte
   Dim RdoMat As ADODB.Recordset

   Dim iOldLevel As Integer
   iOldLevel = Val(lblTyp)
   cmbRun.Clear
   sPartNumber = Compress(cmbPrt)
   FindPart cmbPrt
   lblTyp = iOldLevel
   On Error GoTo DiaErr1
   AdoQry.Parameters(0).Value = Compress(cmbPrt)
   bSqlRows = clsADOCon.GetQuerySet(RdoMat, AdoQry)
   If bSqlRows Then
      With RdoMat
         cmbRun = Format(!Runno, "####0")
         Do Until .EOF
            AddComboStr cmbRun.hWnd, Format$(!Runno, "####0")
            .MoveNext
         Loop
         ClearResultSet RdoMat
      End With
      GetRuns = 1
   Else
      sPartNumber = ""
      GetRuns = 0
   End If
   If cmbRun.ListCount > 0 Then GetThisRun
   On Error Resume Next
   Set RdoMat = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getruns"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub FillMaterial()
   'Dim RdoMat As ADODB.recordset
   On Error GoTo DiaErr1
   cmbPpr.Clear
'   sSql = "SELECT PARTREF,PARTNUM,PALEVEL,PATOOL FROM " _
'          & "PartTable WHERE (PALEVEL<6 AND PATOOL=0) AND " _
'          & "PARTREF<>'" & Compress(cmbPrt) & "' ORDER BY PARTREF"

   'just show parts with open MO's
   sSql = "SELECT DISTINCT PARTREF,PARTNUM" & vbCrLf _
      & "FROM PartTable" & vbCrLf _
      & "WHERE PALEVEL < 6 AND PATOOL = 0" & vbCrLf _
      & "AND PARTREF <> '" & Compress(cmbPrt) & "' AND PAOBSOLETE = 0 ORDER BY PARTREF"
   LoadComboBox cmbPpr
   If cmbPpr.ListCount > 0 Then cmbPpr = cmbPpr.List(0)
   Exit Sub
   
DiaErr1:
   sProcName = "fillmater"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub lblDsc_Change()
   If Left(lblDsc, 8) = "*** Part" Then
      lblDsc.ForeColor = ES_RED
   Else
      lblDsc.ForeColor = vbBlack
   End If
   
End Sub

Private Sub optPick_Click()
   '    If optPick.Value = vbUnchecked Then
   '        optComplete.Value = vbUnchecked
   '        optComplete.Enabled = False
   '    Else
   '        optComplete.Enabled = True
   '    End If
   '
End Sub

Private Sub txtCmt_LostFocus()
   txtCmt = CheckLen(txtCmt, 2048)
   txtCmt = StrCase(txtCmt, ES_FIRSTWORD)
   
End Sub


Private Sub txtQty_LostFocus()
   txtQty = CheckLen(txtQty, 9)
   txtQty = Format(Abs(Val(txtQty)), ES_QuantityDataFormat)
   
End Sub



Private Function FindMatPart() As Byte
   Dim RdoMat As ADODB.Recordset
   
   On Error GoTo DiaErr1
   sSql = "SELECT PARTREF,PARTNUM,PADESC,PAUNITS,PALEVEL,PASTDCOST," _
          & "PAQOH,PALOTTRACK,PATOOL,PAUSEACTUALCOST FROM PartTable WHERE  " _
          & "(PARTREF='" & Compress(cmbPpr) & "' AND PATOOL=0)"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoMat, ES_FORWARD)
   If bSqlRows Then
      With RdoMat
         cmbPpr = "" & Trim(!PartNum)
         lblPsc = "" & Trim(!PADESC)
         lblUom = "" & !PAUNITS
         lblCst = Format(!PASTDCOST, ES_QuantityDataFormat)
         lblQty = Format(!PAQOH, ES_QuantityDataFormat)
         lblTyp = Format(0 + !PALEVEL, "0")
         optLot.Value = !PALOTTRACK
         optActualCost.Value = !PAUSEACTUALCOST
      End With
      cmdChg.Enabled = True
      FindMatPart = 1
   Else
      cmdChg.Enabled = False
      FindMatPart = 0
   End If
   '    optComplete.Value = vbUnchecked
   On Error Resume Next
   Set RdoMat = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "findmatpa"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function



Private Sub PickAddedPart()
   Dim bResponse As Byte
   Dim bLotsRqd As Byte
   Dim lotsAllocated As Boolean
   
   Dim I As Integer
   Dim iLength As Integer
   Dim iLots As Integer
   Dim iPkRecord As Integer
   
   Dim nextINNUMBER As Long
   Dim nextLOIRECORDForLot As Long
   Dim lowestINNUMBER As Long
   
   Dim cLotQty As Currency
   Dim cPckQty As Currency
   Dim cItmLot As Currency
   Dim cQuantity As Currency
   Dim cRemPQty As Currency
   
   Dim sDate As String
   Dim sLot As String
   Dim sMsg As String
   Dim sMoRun As String * 9
   Dim sMoPart As String * 31
   Dim sNewPart As String
   Dim sStatus As String
   
   On Error GoTo whoops
   
   sDate = Format(ES_SYSDATE, "mm/dd/yy")
   If Val(txtQty) = 0 Then
      MsgBox "You Have Entered I Zero Quantity.", vbInformation, Caption
      txtQty.SetFocus
      Exit Sub
   End If
      
   bLotsRqd = CheckLotStatus()
   If bLotsRqd = 1 And optLot.Value = 1 Then
      If Val(txtQty) > Val(lblQty) Then
         MsgBox "This Part Number Is Lot Tracked And There" & vbCrLf _
            & "Aren't Enough On Hand To Satisfy The Need.", _
            vbInformation, Caption
         Exit Sub
      End If
   End If
   sMsg = "You Have Chosen To Add " & txtQty & " " & lblUom & vbCrLf _
          & "Part Number " & cmbPpr & " To The Pick List." & vbCrLf _
          & "Do You Wish To Continue?"
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      'MouseCursor ccHourglass
      cQuantity = Format(Val(txtQty), ES_QuantityDataFormat)
      sNewPart = Compress(cmbPpr)
      cmdChg.Enabled = False
      iLength = Len(Trim(str(cmbRun)))
      iLength = 5 - iLength
      sPartNumber = Compress(cmbPrt)
      sMoPart = cmbPrt
      sMoRun = "RUN" & Space$(iLength) & cmbRun
      bResponse = GetPartAccounts(sNewPart, sDebitAcct, sCreditAcct)

         
      'user lot selection
      Erase lots
      Es_TotalLots = 0
         
'         iLots = GetPartLots(Compress(sNewPart))
      If bLotsRqd = 1 And optLot.Value = vbChecked Then
         'Reqd and Get The lots
         LotSelect.lblPart = sNewPart     'Trim(cmbPpr)
         LotSelect.lblRequired = Abs(cQuantity)
         LotSelect.Show vbModal
         If Es_TotalLots > 0 Then


            lotsAllocated = True
            clsADOCon.BeginTrans
            clsADOCon.ADOErrNum = 0
            
         Else
            lotsAllocated = False
         End If
         
      'automatic lot selection
      Else
         clsADOCon.BeginTrans
         clsADOCon.ADOErrNum = 0
         
         Dim lot As New ClassLot
         lotsAllocated = lot.AutoAllocateLots(sNewPart, cQuantity)
               
      End If
         
      If Not lotsAllocated Then
         clsADOCon.RollbackTrans
         MsgBox "Insufficient lot quantity or lots not allocated by user." & vbCrLf _
            & "Unable to proceed."
         Exit Sub
      End If
      
''''''''''''''''''''''''''''''''''''''''''''''
         
      'create new pick item record
      MouseCursor ccHourglass
      nextINNUMBER = GetLastActivity()
      lowestINNUMBER = nextINNUMBER + 1
      iPkRecord = GetNextPickRecord(sPartNumber, Val(cmbRun))
      sSql = "INSERT INTO MopkTable (PKPARTREF,PKMOPART,PKMORUN," _
         & "PKTYPE,PKPDATE,PKADATE,PKPQTY,PKAQTY,PKCOMT,PKRECORD,PKUNITS) " & vbCrLf _
         & "VALUES('" & sNewPart & "','" & sPartNumber & "'," & cmbRun & ",10,'" & sDate & "','" & sDate & "'," & vbCrLf _
         & cQuantity & "," & cQuantity & ",'" & Trim(txtCmt) & "'," & iPkRecord & "," & "'" & lblUom & "')"
      clsADOCon.ExecuteSQL sSql
      
      'create an inventory activity and a LoitTable record for each
      If Es_TotalLots > 0 Then
         For I = 1 To UBound(lots)
            
            ' Get the Standard and Actual Cost options
            ' 7/3/2009
            Dim strCost As String
            If (optActualCost.Value) Then
                strCost = Format(lots(I).LotCost, ES_QuantityDataFormat)
            Else
                strCost = lblCst
            End If
            
            nextINNUMBER = nextINNUMBER + 1
            sSql = "INSERT INTO InvaTable (INTYPE,INPART,INREF1,INREF2,INPQTY,INAQTY," & vbCrLf _
               & "INAMT,INCREDITACCT,INDEBITACCT,INMOPART,INMORUN,INNUMBER,INLOTNUMBER,INUSER)" & vbCrLf _
               & "VALUES(10,'" & sNewPart & "','PICK','" & Trim(sMoPart) & " " & sMoRun & "'," & vbCrLf _
               & "-" & lots(I).LotSelQty & ",-" & lots(I).LotSelQty & "," & Val(strCost) & "," & vbCrLf _
               & "'" & sCreditAcct & "','" & sDebitAcct & "','" & sPartNumber & "'," & vbCrLf _
               & Val(cmbRun) & "," & nextINNUMBER & ",'" & lots(I).LotSysId & "','" & sInitials & "')"
            clsADOCon.ExecuteSQL sSql
            
            nextLOIRECORDForLot = GetNextLotRecord(lots(I).LotSysId)
            sSql = "INSERT INTO LoitTable (LOINUMBER,LOIRECORD,LOITYPE,LOIPARTREF,LOIQUANTITY," & vbCrLf _
               & "LOIMOPARTREF,LOIMORUNNO,LOIACTIVITY,LOICOMMENT)" & vbCrLf _
               & "VALUES('" & lots(I).LotSysId & "'," & nextLOIRECORDForLot & ",10," & vbCrLf _
               & "'" & sNewPart & "',-" & lots(I).LotSelQty & ",'" & sPartNumber & "'," & vbCrLf _
               & Val(cmbRun) & "," & nextINNUMBER & ",'Picked Item')"
            clsADOCon.ExecuteSQL sSql
            
            sSql = "UPDATE LohdTable SET LOTREMAININGQTY = LOTREMAININGQTY" & " - " & lots(I).LotSelQty & vbCrLf _
               & "WHERE LOTNUMBER = '" & lots(I).LotSysId & "'"
            clsADOCon.ExecuteSQL sSql
         Next
      End If
      
      'update part QOH
      sSql = "UPDATE PartTable SET PAQOH = PAQOH - " & Abs(cQuantity) & "," & vbCrLf _
         & "PALOTQTYREMAINING = PALOTQTYREMAINING - " & Abs(cQuantity) & vbCrLf _
         & "WHERE PARTREF  ='" & sNewPart & "' "
      clsADOCon.ExecuteSQL sSql
      sStatus = GetStatus
      lblStatus = sStatus
      sSql = "UPDATE RunsTable SET RUNSTATUS='" & sStatus & "' WHERE RUNREF='" _
             & sPartNumber & "' AND RUNNO=" & Val(cmbRun) & " "
      clsADOCon.ExecuteSQL sSql
         
      AverageCost sNewPart
      UpdateWipColumns lowestINNUMBER
      
      clsADOCon.CommitTrans
      MouseCursor ccDefault
      
      MsgBox "Material Added And Picked.", vbInformation, Caption
      txtQty = ""
      lblQty = ""
      lblUom = ""
      lblCst = ""
      txtCmt = ""
      If Left(MdiSect.Caption, 3) = "Pro" Then
         If optPrompt.Value = vbChecked Then
            sMsg = "Yes, I Want To Enter An MO For This Item" & vbCrLf _
                   & "No, I Want To Enter I PO For This Item" & vbCrLf _
                   & "Cancel And Let Me Continue Working Here."
            bResponse = MsgBox(sMsg, vbYesNoCancel, Caption)
            If bResponse = vbYes Then
               ShopSHe01a.cmbPrt = cmbPpr
               ShopSHe01a.Show
               Unload Me
            ElseIf bResponse = vbNo Then
               PurcPRe01a.Show
               Unload Me
            Else
               cmbRun.SetFocus
            End If
         Else
            cmbRun.SetFocus
         End If
      Else
         cmbRun.SetFocus
      End If
   End If
   Exit Sub
   
whoops:
   ProcessError "PickAddedPart"
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiProd", "mpadd", sOptions)
   If Len(sOptions) Then
      optPrompt.Value = Val(sOptions)
   Else
      optPrompt.Value = vbChecked
   End If
   sOptions = GetSetting("Esi2000", "EsiProd", "mpaddauto", sOptions)
   If Len(sOptions) Then
      optPick.Value = Val(sOptions)
   Else
      optPick.Value = vbChecked
   End If
   
End Sub

Private Sub SaveOptions()
   SaveSetting "Esi2000", "EsiProd", "mpadd", optPrompt.Value
   SaveSetting "Esi2000", "EsiProd", "mpaddauto", optPick.Value
   
End Sub

Private Sub AddPickItem()
   Dim iPkRecord As Integer
   Dim cQuantity As Currency
   Dim sNewPart As String
   Dim sStatus As String
   
   Dim sDate As Variant
   sDate = Format(ES_SYSDATE, "mm/dd/yy")
   sNewPart = Compress(cmbPpr)
   iPkRecord = GetNextPickRecord(sPartNumber, Val(cmbRun))
   
'   Select Case lblStatus
'      Case "SC", "RL"
'         sStatus = "PL"
'      Case Else
'         sStatus = "PP"
'   End Select
   cQuantity = Format(Val(txtQty), ES_QuantityDataFormat)
   
   clsADOCon.ADOErrNum = 0
   clsADOCon.BeginTrans
   
   
   sSql = "INSERT INTO MopkTable (PKPARTREF,PKMOPART,PKMORUN," _
          & "PKTYPE,PKPDATE,PKADATE,PKPQTY,PKAQTY,PKCOMT,PKRECORD,PKUNITS) " _
          & "VALUES('" & sNewPart & "','" & sPartNumber & "'," _
          & Val(cmbRun) & ",9,'" & sDate & "','" & sDate & "'," _
          & cQuantity & ",0,'" & txtCmt & "'," & iPkRecord _
          & ",'" & lblUom & "')"
   clsADOCon.ExecuteSQL sSql
   
   Dim mo As New ClassMO
   sStatus = mo.GetOpenMoStatus(sPartNumber, cmbRun)
   
   sSql = "UPDATE RunsTable SET RUNSTATUS='" & sStatus & "' WHERE RUNREF='" _
          & sPartNumber & "' AND RUNNO=" & Val(cmbRun) & " "
   clsADOCon.ExecuteSQL sSql
   If clsADOCon.ADOErrNum = 0 Then
      clsADOCon.CommitTrans
      lblStatus = sStatus
      cmbPpr = ""
      txtQty = ""
      lblQty = ""
      lblUom = ""
      lblCst = ""
      txtCmt = ""
      SysMsg "Pick Item Was Added.", True
   Else
      clsADOCon.RollbackTrans
      clsADOCon.ADOErrNum = 0
      
      MsgBox "Could Not Add This Item.", vbInformation, Caption
   End If
   
End Sub

'Private Function GetStatus() As String
'   Dim RdoPck As ADODB.recordset
'   'On Error Resume Next
'   sSql = "SELECT PKMOPART,PKMORUN,PKAQTY,PKTYPE FROM MopkTable" & vbCrLf _
'          & "WHERE (PKAQTY=0 AND PKTYPE=9 AND PKMOPART='" & sPartNumber & "' AND PKMORUN=" & Val(cmbRun) & ")"
'   bsqlrows = clsadocon.getdataset(ssql, RdoPck, ES_FORWARD)
'   If bSqlRows Then
'      With RdoPck
'         GetStatus = "PP"
'         ClearResultSet RdoPck
'      End With
'   Else
'      GetStatus = "PC"
'   End If
'
'End Function
'

Private Function GetStatus() As String
   Dim mo As New ClassMO
   GetStatus = mo.GetOpenMoStatus(sPartNumber, cmbRun)
End Function



Private Function IsOpenPartAlreadyOnPickList(sNewPart As String) As Boolean
   'determine whether there is already a pick list item for this part
   
   Dim Ado As ADODB.Recordset
   sSql = "select count(*) as ct FROM MopkTable" & vbCrLf _
          & "where PKMOPART='" & sPartNumber & "' AND PKMORUN=" & Val(cmbRun) & vbCrLf _
          & "and PKPARTREF='" & Compress(sNewPart) & "' AND (PKTYPE=9 or PKTYPE=23)"
   IsOpenPartAlreadyOnPickList = False
   If clsADOCon.GetDataSet(sSql, Ado, ES_FORWARD) Then    '<> 0 Then 'is a row returned
      If Ado!ct > 0 Then
         IsOpenPartAlreadyOnPickList = True
      End If
   End If
   Set Ado = Nothing
End Function

