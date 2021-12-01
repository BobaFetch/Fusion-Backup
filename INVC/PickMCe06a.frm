VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form PickMCe06a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Auto Pick Items for MO"
   ClientHeight    =   2475
   ClientLeft      =   2475
   ClientTop       =   645
   ClientWidth     =   6345
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   6345
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "PickMCe06a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.TextBox txtTyp 
      Height          =   285
      Left            =   5400
      MaxLength       =   1
      TabIndex        =   13
      Tag             =   "1"
      Top             =   1440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.OptionButton optInd 
      Caption         =   "&Individual  "
      Height          =   255
      Left            =   1320
      TabIndex        =   11
      Top             =   1920
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.OptionButton optExp 
      Caption         =   "&Exceptions  "
      Height          =   255
      Left            =   2640
      TabIndex        =   10
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CheckBox optItm 
      Caption         =   "picks"
      Height          =   255
      Left            =   5160
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1920
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdPck 
      Caption         =   "&Pick Items"
      Height          =   315
      Left            =   5400
      TabIndex        =   2
      ToolTipText     =   "Pick Items"
      Top             =   480
      Width           =   915
   End
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   1320
      TabIndex        =   0
      ToolTipText     =   "Select Part Number"
      Top             =   720
      Width           =   3545
   End
   Begin VB.ComboBox cmbRun 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1320
      TabIndex        =   1
      ToolTipText     =   "Select Run Number"
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5400
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   915
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   0
      Top             =   2280
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2475
      FormDesignWidth =   6345
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Type"
      Height          =   285
      Index           =   3
      Left            =   4440
      TabIndex        =   14
      Top             =   1440
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Options"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   12
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Run Number"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label lblStat 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3720
      TabIndex        =   6
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1320
      TabIndex        =   5
      Top             =   1080
      Width           =   3135
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Status (PL, PP)"
      Height          =   255
      Index           =   15
      Left            =   2520
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
   End
End
Attribute VB_Name = "PickMCe06a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Stanwood, Washington, USA  ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'12/7/04 changed Lots - See procedure in PickMCe01c
Option Explicit
Dim AdoQry As ADODB.Command
Dim AdoParameter As ADODB.Parameter
Dim bGoodRuns As Byte
Dim bGoodMo As Byte
Public bOnLoad As Byte

Dim iTotalItems As Integer
Dim bFIFO As Byte


Dim sPartNumber As String
Dim sCreditAcct As String
Dim sDebitAcct As String

Dim sLots(50, 2) As String
'0 = Lot Number
'1 = Lot Quantity
'Dim sPartsGroup(250) As String   'TODO: this is redundant - should just use vitems(irow,1)

Dim vItems(250, 14) As Variant
' 0 = loc
Private Const PICK_LOCATION = 0
' 1 = part
Private Const PICK_PARTNUMBER = 1
' 2 = compressed part
Private Const PICK_PARTREF = 2
' 3 = stand cost
Private Const PICK_STANDARDCOST = 3
' 4 = desc
Private Const PICK_DESCRIPTION = 4
' 5 = rev
Private Const PICK_REVISION = 5
' 6 = planned
Private Const PICK_REQUIREDQTY = 6
' 7 = actual
Private Const PICK_QUANTITY = 7
' 8 = wip location
Private Const PICK_WIPLOCATION = 8
' 9 = complete?
Private Const PICK_ITEMCOMPLETE = 9
'10 = Qoh
Private Const PICK_QOH = 10
'11 = LotTracked Part
Private Const PICK_LOTTRACKED = 11
'12 = PKRECORD
Private Const PICK_PKRECORD = 12
'13 = Unit of Measure
Private Const PICK_UNITS = 13


Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'when calling from other forms, place MO info here
Public PassedInMoPartNo As String
Public PassedInMoRunNo As Integer

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
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
   
   If cmbPrt <> "" Then bGoodRuns = GetRuns()
   
End Sub


Private Sub cmbRun_Click()
   bGoodMo = GetPart()
   
End Sub

Private Sub cmbRun_GotFocus()
   cmbRun_Click
   
End Sub


Private Sub cmbRun_LostFocus()
   cmbRun = CheckLen(cmbRun, 5)
   cmbRun = Format(Abs(Val(cmbRun)), "####0")
   bGoodMo = GetPart()
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub


Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   cmbPrt = ""
   
End Sub


Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      If Left(Caption, 3) <> "Rev" Then
         OpenHelpContext "5203"
      Else
         OpenHelpContext "5201"
      End If
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub


Private Sub cmdPck_Click()
   If Not bGoodMo Then
      MsgBox "Run (PL or PP) For This Part Wasn't Found.", vbInformation, Caption
      Exit Sub
   Else
      Dim bGoodItem As Boolean
      
      ' Get all Items
      bGoodItem = GetItems()
   
      If (bGoodItem) Then
         PickItems
      End If
      
   End If
   
End Sub

Function GetItems() As Boolean
   Dim RdoPck As ADODB.Recordset
   Dim bLotsAct As Byte
   Dim iRow As Integer
   
   MouseCursor 13
   sPartNumber = Compress(cmbPrt)
   bLotsAct = CheckLotStatus()
   iRow = 0
   Erase vItems
   'Erase sPartsGroup
   On Error GoTo DiaErr1
   sSql = "SELECT PARTREF,PARTNUM,PADESC,PASTDCOST,PAQOH,PALOCATION," _
          & "PALOTTRACK,PAUNITS,PKPARTREF,PKMOPART,PKMORUN,PKREV,PKTYPE,PKPQTY," _
          & "PKAQTY,PKRECORD,PKUNITS FROM PartTable,MopkTable WHERE PARTREF=PKPARTREF " _
          & "AND PKMOPART='" & sPartNumber & "' AND PKMORUN=" _
          & Trim(Val(cmbRun)) & " AND PKAQTY=0 AND (PKTYPE=9 or PKTYPE=23) " _
          & "ORDER BY PALOCATION,PARTREF"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPck, ES_KEYSET)
   If bSqlRows Then
      With RdoPck
         Do Until .EOF
            iRow = iRow + 1
            If iRow > 299 Then
               MouseCursor 0
               sSql = "This Pick Has More Than 300 Items." & vbCr _
                      & "You Will Have To Pick In (2) Steps."
               MsgBox sSql, vbInformation, Caption
               Exit Do
            End If
            vItems(iRow, PICK_LOCATION) = "" & Trim(!PALOCATION)
            vItems(iRow, PICK_PARTNUMBER) = "" & Trim(!PartRef)
            'spartsgroup(iRow) = "" & Trim(!PartRef)
            vItems(iRow, PICK_PARTREF) = "" & Trim(!PartNum)
            vItems(iRow, PICK_STANDARDCOST) = Format(!PASTDCOST, "#####0.000")
            vItems(iRow, PICK_DESCRIPTION) = "" & Trim(!PADESC)
            vItems(iRow, PICK_REVISION) = "" & Trim(!PKREV)
            vItems(iRow, PICK_REQUIREDQTY) = Format(!PKPQTY, "####0.000")
            vItems(iRow, PICK_QUANTITY) = Format(!PKPQTY, "####0.000")
            vItems(iRow, PICK_ITEMCOMPLETE) = 1
            vItems(iRow, PICK_WIPLOCATION) = ""
            vItems(iRow, PICK_QOH) = Format(!PAQOH, "####0.000")
            If bLotsAct = 1 Then
               vItems(iRow, PICK_LOTTRACKED) = Format(!PALOTTRACK, "0")
            Else
               vItems(iRow, PICK_LOTTRACKED) = 0
            End If
            vItems(iRow, PICK_PKRECORD) = Format(!PKRECORD)
            '5/13/04
            If Trim(!PKUNITS) = "" Then
               vItems(iRow, PICK_UNITS) = Format(!PAUNITS)
            Else
               vItems(iRow, PICK_UNITS) = Format(!PKUNITS)
            End If
            .MoveNext
         Loop
         ClearResultSet RdoPck
      End With
      iTotalItems = iRow
   End If

   GetItems = True
'   sSql = "UPDATE RunsTable SET RUNSTATUS='PC'" & vbCrLf _
'      & "WHERE RUNREF='" & Compress(cmbPrt) & "' AND RUNNO=" & cmbRun
'   clsADOCon.ExecuteSQL sSql
   
   
   Set RdoPck = Nothing
   MouseCursor 0
   Exit Function
   
DiaErr1:
   sProcName = "getitems"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
   GetItems = False

   
End Function



Private Sub PickItems()
   Dim RdoPck As ADODB.Recordset
   Dim iRow As Integer
   Dim I As Integer
   Dim iLotsAvail As Integer
   Dim iPkRecord As Integer
   
   Dim bBadPick As Byte
   Dim bGoodPick As Byte
   Dim bResponse As Byte
   
   'lots
   Dim bLotsAct As Byte
   Dim bPartLot As Byte
   Dim bLotFail As Byte
   Dim bLotsFailed As Byte
   Dim bItemsPicked As Byte
   
   Dim iLots As Integer
   Dim iRef As Integer
   Dim iTrans As Integer
   
   Dim lCOUNTER As Long
   Dim lSysCount As Long
   Dim lLOTRECORD As Long
   
   Dim cCost As Currency
   Dim clineCost As Currency
   Dim lotQty As Currency
   Dim PickQty As Currency
   'Dim cItmLot As Currency
   Dim cQuantity As Currency
   Dim remainingPickQty As Currency
   
   'Costs
   Dim cMaterial As Currency
   Dim cLabor As Currency
   Dim cExpense As Currency
   Dim cOverhead As Currency
   Dim cHours As Currency
   
   Dim sLot As String
   Dim sMsg As String
   
   Dim MoPartNumber As String
   Dim moPartRef As String
   Dim moRunNo As Long
   Dim sMoRun As String * 9
   
   Dim sNewDate As String
   Dim sNewPart As String
   Dim sNewRev As String
   Dim sPickPart As String
   Dim sComment As String
   Dim vAdate As Variant
   
   MoPartNumber = Trim(cmbPrt)
   moPartRef = Compress(cmbPrt)
   moRunNo = Val(cmbRun)
   'sMoPart = Compress(cmbPrt)
   vAdate = Format(GetServerDateTime(), "mm/dd/yy hh:mm")
   bLotsAct = CheckLotStatus()
   
   'verify that there are sufficient lot quantities for all lot-tracked items
   Dim msg As String, failures As Integer
   msg = "The following parts have insufficient lot quantities:" & vbCrLf
   failures = 0
   If bLotsAct = 1 Then
      For iRow = 1 To iTotalItems
         If Val(vItems(iRow, PICK_LOTTRACKED)) = 1 And Val(vItems(iRow, PICK_QUANTITY)) > 0 Then
            cQuantity = CCur(vItems(iRow, PICK_QUANTITY))
            'lotQty = GetRemainingLotQty(CStr(vItems(iRow, PICK_PARTNUMBER)))
            lotQty = GetLotRemainingQty(CStr(vItems(iRow, PICK_PARTNUMBER)))
            If lotQty < cQuantity Then
               failures = failures + 1
               msg = msg & "Part " & CStr(vItems(iRow, PICK_PARTNUMBER)) & " has only " & lotQty & " available.  You have requested " & cQuantity & vbCrLf
            End If
         End If
      Next
      If failures > 0 Then
         msg = msg & "Please make corrections before continuing."
         MsgBox msg
         Exit Sub
      End If
   Else
       For iRow = 1 To iTotalItems
         If Val(vItems(iRow, PICK_QUANTITY)) > 0 Then
            cQuantity = CCur(vItems(iRow, PICK_QUANTITY))
            lotQty = CCur(vItems(iRow, PICK_QOH))
            If lotQty < cQuantity Then
               failures = failures + 1
               msg = msg & "Part " & CStr(vItems(iRow, PICK_PARTNUMBER)) & " has only " & lotQty & " available.  You have requested " & cQuantity & vbCrLf
            End If
         End If
      Next
      If failures > 0 Then
         msg = msg & "Please make corrections before continuing."
         MsgBox msg
         Exit Sub
      End If
   End If
   
   'remove prior lot selections from temporary table
   sSql = "delete from TempPickLots" & vbCrLf _
      & "where ( MoPartRef = '" & moPartRef & "' and MoRunNo = " & cmbRun & " )" & vbCrLf _
      & "or DateDiff( hour, WhenCreated, getdate() ) > 24"
   clsADOCon.ExecuteSQL sSql
   
   'Everything seems ok, so let's pick it
   'MouseCursor ccHourglass
   iRow = Len(Trim(str(cmbRun)))
   iRow = 5 - iRow
   sMoRun = "RUN" & Space$(iRow) & cmbRun
'   If sJournalID <> "" Then iTrans = GetNextTransaction(sJournalID)
   lCOUNTER = GetLastActivity()
   lSysCount = lCOUNTER + 1
   bItemsPicked = 0
   
   For iRow = 1 To iTotalItems
      If Val(vItems(iRow, PICK_QUANTITY)) > 0 Then
         
         bPartLot = Val(vItems(iRow, PICK_LOTTRACKED))
         cQuantity = Format(Val(vItems(iRow, PICK_QUANTITY)), "########0.000")
         sPickPart = vItems(iRow, PICK_PARTREF)
         
'         If bPartLot = 1 Then
'            lotQty = GetLotRemainingQty(CStr(vItems(iRow, PICK_PARTNUMBER)))
'         Else
'            lotQty = cQuantity
'         End If
         
         'before beginning transaction, let user select lots
         lotQty = 0
         remainingPickQty = cQuantity
         Es_TotalLots = 0
         Erase lots
         
         
         iLots = GetPartLots(Compress(sPickPart))
         
         For I = 1 To iLots
            lotQty = Val(sLots(I, 1))
            If lotQty >= remainingPickQty Then
               PickQty = remainingPickQty
               lotQty = lotQty - remainingPickQty
               remainingPickQty = 0
            Else
               PickQty = lotQty
               remainingPickQty = remainingPickQty - lotQty
               lotQty = 0
            End If
            
            'save info for lot selection
            sSql = "INSERT INTO TempPickLots ( MoPartRef, MoRunNo, PickPartRef, LotID, LotQty,selIndex )" & vbCrLf _
               & "Values ( '" & moPartRef & "', " & moRunNo & ", '" & vItems(iRow, PICK_PARTNUMBER) & "'," & vbCrLf _
               & "'" & sLots(I, 0) & "', " & PickQty & ", " & iRow & " )"
            clsADOCon.ExecuteSQL sSql
            If remainingPickQty <= 0 Then Exit For
         Next
      
      End If
      
   Next
      
   'begin transaction if not already started
   clsADOCon.ADOErrNum = 0
   clsADOCon.BeginTrans
   
   MouseCursor ccHourglass
   
   'make sure lot selections are still available
   Dim rdo As ADODB.Recordset
   sSql = "select count(*) as ct" & vbCrLf _
      & "from TempPickLots tmp" & vbCrLf _
      & "join LohdTable lot on tmp.LotID = lot.LotNumber" & vbCrLf _
      & "where lot.LotRemainingQty < tmp.LotQty" & vbCrLf _
      & "and tmp.MoPartRef = '" & moPartRef & "' and tmp.MoRunNo = " & moRunNo
   If clsADOCon.GetDataSet(sSql, rdo, ES_FORWARD) Then
      If rdo!ct > 0 Then
         If bLotsAct = 1 Then
            clsADOCon.RollbackTrans
         End If
         MsgBox "Another user has allocated quantities from the lots selected.  Please try again."
         Set rdo = Nothing
         Exit Sub
      End If
   End If
   'rdo.Close
   Set rdo = Nothing
   'pick the items
   Dim pick As New ClassPick
   Dim cUnitCost As Currency
   Dim PickRecordNo As Long
   Dim pickUnitOfMeasure As String, wipLocation As String
   Dim pickRequiredQty As Currency
   Dim pickComplete As Boolean
   
   pick.MoPartNumber = MoPartNumber
   pick.MoRunNumber = moRunNo
   
   For iRow = 1 To iTotalItems
      cQuantity = Val(vItems(iRow, PICK_QUANTITY))
      pickRequiredQty = Val(vItems(iRow, PICK_REQUIREDQTY))
      sPickPart = vItems(iRow, PICK_PARTREF)
      cUnitCost = Val(vItems(iRow, PICK_STANDARDCOST))
      PickRecordNo = CLng(vItems(iRow, PICK_PKRECORD))
      pickUnitOfMeasure = vItems(iRow, PICK_UNITS)
      wipLocation = vItems(iRow, PICK_WIPLOCATION)
      pickComplete = IIf(vItems(iRow, PICK_ITEMCOMPLETE) = 0, False, True)
      If cQuantity > 0 Then
         ' Set the Index for the Picked records
         pick.PickPartIndex = iRow
         
         If Not pick.PickPart(sPickPart, cQuantity, pickRequiredQty, pickUnitOfMeasure, cUnitCost, _
            PickRecordNo, wipLocation, pickComplete, True) Then
            
            clsADOCon.RollbackTrans
            MouseCursor ccDefault
            MsgBox "Could Not Successfully Complete The Pick.", _
               vbExclamation, Caption
            Exit Sub
         End If
            
      End If
   Next
   
   If (clsADOCon.ADOErrNum = 0) Then
      clsADOCon.CommitTrans
      MsgBox "Successfully Complete The Pick.", _
         vbExclamation, Caption
   Else
      clsADOCon.RollbackTrans
      MsgBox "Could Not Successfully Complete The Pick.", _
         vbExclamation, Caption
   End If
   
   MouseCursor ccDefault
   
   Exit Sub
      
DiaErr1:
   If (clsADOCon.ADOErrNum <> 0) Then
     clsADOCon.RollbackTrans
   End If
   sProcName = "PickItems"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub
   

Private Sub Form_Activate()
   Dim b As Byte
   Dim iList As Integer
   
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      'Can't find the caption
      bFIFO = GetInventoryMethod()

      iList = SetRecent(Me)
      If sPassedMo <> "" Then cmbPrt = sPassedMo
      sJournalID = GetOpenJournal("IJ", Format$(ES_SYSDATE, "mm/dd/yy"))
      If Left(sJournalID, 4) = "None" Then
         sJournalID = ""
         b = 1
      Else
         If sJournalID = "" Then b = 0 Else b = 1
      End If
      If b = 0 Then
         MouseCursor 0
         MsgBox "There Is No Open Inventory Journal For This Period.", _
            vbExclamation, Caption
         Sleep 500
         Unload Me
         Exit Sub
      End If
      FillCombo
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST
   FormatControls
   sSql = "SELECT RUNREF,RUNSTATUS,RUNNO FROM " _
          & "RunsTable WHERE RUNREF = ? " _
          & "AND (RUNSTATUS='PL' OR RUNSTATUS='PP')"
   'Set RdoQry = RdoCon.CreateQuery("", sSql)
   Set AdoQry = New ADODB.Command
   AdoQry.CommandText = sSql
   
   Set AdoParameter = New ADODB.Parameter
   AdoParameter.Size = 30
   AdoParameter.Type = adChar
   
   AdoQry.Parameters.Append AdoParameter
   bOnLoad = 1
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If optItm = vbChecked Then
      If Left(Caption, 3) = "Rev" Then
         Unload PickMCe01b
      Else
         Unload PickMCe01c
      End If
   End If
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   SaveCurrentSelections
   FormUnload
   Set AdoParameter = Nothing
   Set AdoQry = Nothing
   Set PickMCe06a = Nothing
   
End Sub



Public Sub FillCombo()
   On Error GoTo DiaErr1
   cmbPrt.Clear
   sSql = "SELECT DISTINCT PARTREF,PARTNUM,PADESC,RUNREF " _
          & "FROM PartTable,RunsTable WHERE PARTREF=RUNREF " _
          & "AND (RUNSTATUS='PL' OR RUNSTATUS='PP') ORDER BY PARTREF"
   LoadComboBox cmbPrt
   If cmbPrt.ListCount > 0 Then
      cmbPrt = cmbPrt.List(0)
      GetPart
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Function GetRuns() As Byte
   Dim RdoRns As ADODB.Recordset
   
   cmbRun.Clear
   On Error GoTo DiaErr1
   AdoQry.Parameters(0).Value = Compress(cmbPrt)
   bSqlRows = clsADOCon.GetQuerySet(RdoRns, AdoQry)
   If bSqlRows Then
      With RdoRns
         cmbRun = Format(!Runno, "####0")
         lblStat = "" & !RUNSTATUS
         Do Until .EOF
            AddComboStr cmbRun.hWnd, Format$(!Runno, "####0")
            .MoveNext
         Loop
         ClearResultSet RdoRns
      End With
      GetRuns = True
   Else
      GetRuns = False
   End If
   If GetRuns Then bGoodMo = GetPart()
   Set RdoRns = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getruns"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Function GetPart() As Byte
   Dim RdoPrt As ADODB.Recordset
   
   On Error GoTo DiaErr1
   sSql = "SELECT DISTINCT PARTREF,PARTNUM,PADESC,PALEVEL,RUNREF,RUNSTATUS " _
          & "FROM PartTable,RunsTable WHERE PARTREF=RUNREF " _
          & "AND PARTREF='" & Compress(cmbPrt) & "' AND RUNNO=" & str(Val(cmbRun)) & " " _
          & "AND (RUNSTATUS='PL' OR RUNSTATUS='PP')"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPrt)
   If bSqlRows Then
      With RdoPrt
         cmbPrt = "" & Trim(!PartNum)
         lblDsc = "" & Trim(!PADESC)
         lblStat = "" & Trim(!RUNSTATUS)
         txtTyp = Format(!PALEVEL, "0")
         cUR.CurrentPart = cmbPrt
         ClearResultSet RdoPrt
         GetPart = True
      End With
   Else
      GetPart = False
   End If
   Set RdoPrt = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getpart"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub optExp_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optInd_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optItm_Click()
   'never visible. marks PickMCe01c as loaded
   
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
            sLots(iList, 1) = Format$(!LOTREMAININGQTY, "#####0.000")
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

