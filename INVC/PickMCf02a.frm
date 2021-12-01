VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form PickMCf02a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cancel A Pick List Item"
   ClientHeight    =   5175
   ClientLeft      =   1845
   ClientTop       =   1065
   ClientWidth     =   7950
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   7950
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "PickMCf02a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   23
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CheckBox optDel 
      Alignment       =   1  'Right Justify
      Caption         =   "Delete The Item From The List?"
      Height          =   255
      Left            =   4920
      TabIndex        =   5
      ToolTipText     =   "Delete The Pick Item Completely"
      Top             =   1560
      Value           =   1  'Checked
      Width           =   2895
   End
   Begin VB.CommandButton cmdSel 
      Caption         =   "S&elect"
      Height          =   315
      Left            =   7000
      TabIndex        =   2
      ToolTipText     =   "Get Items"
      Top             =   720
      Width           =   875
   End
   Begin VB.ComboBox cmbPck 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1320
      TabIndex        =   3
      Tag             =   "3"
      ToolTipText     =   "Select From List (No Edit)"
      Top             =   1560
      Width           =   3545
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   7560
      Top             =   4800
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   5175
      FormDesignWidth =   7950
   End
   Begin VB.CommandButton cmdCpl 
      Caption         =   "&Apply"
      Enabled         =   0   'False
      Height          =   315
      Left            =   7000
      TabIndex        =   6
      ToolTipText     =   "Cancel This Pick List Item"
      Top             =   2520
      Width           =   875
   End
   Begin VB.ComboBox cmbRun 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   5880
      TabIndex        =   1
      Tag             =   "1"
      ToolTipText     =   "Select Run Number"
      Top             =   720
      Width           =   855
   End
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   1320
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Select MO Part Number"
      Top             =   720
      Width           =   3545
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   7000
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   2295
      Left            =   1320
      TabIndex        =   4
      ToolTipText     =   "Click The Item Or Scroll And Press Enter"
      Top             =   2640
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   4048
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      AllowBigSelection=   0   'False
      Enabled         =   0   'False
      FocusRect       =   2
      HighLight       =   0
      ScrollBars      =   2
      SelectionMode   =   1
   End
   Begin VB.Label lblItems 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   285
      Left            =   6840
      TabIndex        =   22
      ToolTipText     =   "Selected Items"
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label lblUom 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2200
      TabIndex        =   21
      ToolTipText     =   "Unit Of Measure"
      Top             =   2280
      Width           =   390
   End
   Begin VB.Label Z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      Height          =   255
      Index           =   6
      Left            =   4320
      TabIndex        =   20
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label lblDate 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   4920
      TabIndex        =   19
      ToolTipText     =   "User Date"
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity"
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   18
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label Z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Item"
      Height          =   255
      Index           =   4
      Left            =   2880
      TabIndex        =   17
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label lblQty 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1320
      TabIndex        =   16
      ToolTipText     =   "Quantity"
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   1320
      TabIndex        =   15
      Top             =   1920
      Width           =   3255
   End
   Begin VB.Label lblItm 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3480
      TabIndex        =   14
      ToolTipText     =   "Item Number"
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label Z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Pick Item"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   13
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      Height          =   255
      Index           =   2
      Left            =   5040
      TabIndex        =   12
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label lblStu 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   6240
      TabIndex        =   11
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   0
      Left            =   1320
      TabIndex        =   10
      Top             =   1080
      Width           =   3255
   End
   Begin VB.Label Z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Run"
      Height          =   255
      Index           =   1
      Left            =   5040
      TabIndex        =   9
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   8
      Top             =   720
      Width           =   1095
   End
End
Attribute VB_Name = "PickMCf02a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Stanwood, Washington, USA  ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'10/23/03 New
'10/7/04  See GetOpenLots
'12/23/04 Changed Picks to allow Projects
'1/28/05 Changed to show Picked only, correctly show quantity and added date
'        Added the Grid
'2/2/05 Click and KeyPress to Grid selection
'2/7/05 Added GetOldCost
'10/19/06 Fixed cmbRun (not filling)
Option Explicit

Dim bCancel As Byte
Dim bOnLoad As Byte
Dim bGoodRuns As Byte
Dim bGoodPick As Byte


Dim sPartNumber As String
Dim sCreditAcct As String
Dim sDebitAcct As String
Dim sOldPart As String

Dim vItems(500, 8) As Variant
'0 = Part Number
'1 = Description
'2 = MoNumber
'3 = Run
'4 = Record
'5 = Quantity
'6 = Date
'7 = PKUNITS (Uom)
Dim sPartsGroup(500) As String
Dim sLots(30, 6) As String 'See GetOpenLots

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub SaveOptions()
   Dim sOptions As String
   sOptions = optDel.Value
   SaveSetting "Esi2000", "EsiInvc", "mcnpi", sOptions
   
End Sub


Private Sub GetOptions()
   On Error Resume Next
   optDel.Value = GetSetting("Esi2000", "EsiInvc", "mcnpi", optDel.Value)
   
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub



Private Sub cmbPck_Click()
   On Error Resume Next
   lblDsc(1) = vItems(cmbPck.ListIndex, 1)
   lblItm = vItems(cmbPck.ListIndex, 4)
   lblQty = Format(vItems(cmbPck.ListIndex, 5), ES_QuantityDataFormat)
   lblDate = vItems(cmbPck.ListIndex, 6)
   lblUom = vItems(cmbPck.ListIndex, 7)
   
End Sub


Private Sub cmbPck_LostFocus()
   Dim b As Byte
   Dim item As Integer
   cmbPck = CheckLen(cmbPck, 30)
   
   For item = 0 To cmbPck.ListCount - 1
      If cmbPck = cmbPck.List(item) Then b = 1
   Next
   If b = 0 Then
      Beep
      lblDsc(1) = "Item Wasn't Found."
      lblItm = ""
      lblQty = "0.000"
      bGoodPick = 0
   Else
      On Error Resume Next
      lblDsc(1) = vItems(cmbPck.ListIndex, 1)
      lblItm = vItems(cmbPck.ListIndex, 4)
      lblQty = Format(vItems(cmbPck.ListIndex, 5), ES_QuantityDataFormat)
      lblDate = vItems(cmbPck.ListIndex, 6)
      lblUom = vItems(cmbPck.ListIndex, 7)
      bGoodPick = 1
   End If
   
End Sub


Private Sub cmbPrt_Click()
   bGoodRuns = GetRuns()
   
End Sub


Private Sub cmbPrt_LostFocus()
   cmbPrt = CheckLen(cmbPrt, 30)
   If bCancel Then Exit Sub
   
   If (Not ValidPartNumber(cmbPrt.Text)) Then
      MsgBox "Can't Select The Part Number Which Is Obsolete or Inactive. ", _
         vbInformation, Caption
      cmbPrt = ""
      Exit Sub
   End If
   
   cmdSel.Enabled = True
   bGoodRuns = GetRuns()
   
End Sub


Private Sub cmbRun_Click()
   GetStatus
   
End Sub


Private Sub cmbRun_LostFocus()
   GetStatus
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub



Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   bCancel = True
   
End Sub


Private Sub cmdCpl_Click()
   Dim bResponse As Byte
   Dim sMsg As String
   
   If cmbPck.ListCount = 0 Or Trim(cmbPck) = "" Then
      MsgBox "No Valid Items Selected To Cancel.", _
         vbInformation, Caption
      Exit Sub
   End If
   
   If Val(lblQty) = 0 And optDel.Value = vbUnchecked Then
      MsgBox "The Quantity Is Zero And It Hasn't Been Selected For" & vbCr _
         & "Deletion. The Function Would Not Return Results.", _
         vbInformation, Caption
      Exit Sub
   Else
      If optDel.Value = vbChecked Then
         bResponse = MsgBox("This Function Removes The Item From The Pick List." _
                     & vbCr & "Do You Wish To Continue?", ES_YESQUESTION, Caption)
         If bResponse = vbNo Then
            CancelTrans
            Exit Sub
         End If
      End If
   End If
   If bGoodPick = 0 Then
      MsgBox "That Part Number Wasn't Listed.", vbInformation, Caption
      Exit Sub
   End If
   sMsg = "Cancels The Selected Picked Item And Returns The Items To " & vbCr _
          & "Inventory. Are You Sure That You Want To Continue?"
   bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
   If bResponse = vbYes Then
      On Error GoTo McnplCn1
      cmdCpl.Enabled = False
      bGoodPick = CancelPick()
      If bGoodPick Then
         MouseCursor 0
         MsgBox "Pick List Item Was Canceled.", _
            vbInformation, Caption
         'FillCombo
         GetItems
      Else
         MouseCursor 0
         MsgBox "Could Not Cancel The Pick List Item..", _
            vbInformation, Caption
      End If
   Else
      CancelTrans
   End If
   Exit Sub
   
McnplCn1:
   Resume McnplCn2
   CurrError.Description = Err.Description
McnplCn2:
   MouseCursor 0
   On Error Resume Next
   'RdoCon.RollbackTrans
   sMsg = CurrError.Description & vbCr _
          & "Could Not Complete Pick List Cancel."
   MsgBox sMsg, vbExclamation, Caption
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext "5251"
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub


Private Sub cmdSel_Click()

   Dim sMsg As String
   If (lblStu.Caption = "CL" Or lblStu.Caption = "CO") Then
        sMsg = "MO is completed or closed."
        MsgBox sMsg, vbInformation, Caption
        Exit Sub
   End If
   GetItems
   
End Sub


Private Sub Form_Activate()
   Dim b As Byte
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then FillCombo
   bOnLoad = 0
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   GetOptions
   With Grid1
      .Rows = 2
      .ColWidth(0) = 450
      .ColWidth(1) = 2450
      .ColWidth(2) = 850
      .ColWidth(3) = 1020
      .ColAlignment(1) = 0
      .row = 0
      .Col = 0
      .Text = "Item"
      .Col = 1
      .Text = "Part Number"
      .Col = 2
      .Text = "Date"
      .Col = 3
      .Text = "Quantity"
   End With

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

   Set PickMCf02a = Nothing
   
End Sub



Private Sub FillCombo(Optional SkipGl As Boolean)
   Dim RdoPcl As ADODB.Recordset
   
   Dim b As Byte
   
   On Error GoTo DiaErr1
   If Not SkipGl Then
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
   End If
   cmdCpl.Enabled = False
   cmbPrt.Clear
   Grid1.Rows = 2
   Grid1.row = 1
   Grid1.Col = 0
   Grid1.Text = ""
   Grid1.Col = 1
   Grid1.Text = ""
   Grid1.Col = 2
   Grid1.Text = ""
   Grid1.Col = 3
   Grid1.Text = ""
   sProcName = "fillcombo"
   sSql = "SELECT DISTINCT PKMOPART,PARTREF,PARTNUM,PADESC FROM " _
          & "MopkTable,PartTable WHERE (PKMOPART=PARTREF) AND PKAQTY > 0 " _
          & "AND PKTYPE<12 ORDER BY PKMOPART"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPcl, ES_FORWARD)
   If bSqlRows Then
      With RdoPcl
         cmbPrt = "" & Trim(!PartNum)
         lblDsc(0) = "" & Trim(!PADESC)
         Do Until .EOF
            AddComboStr cmbPrt.hWnd, "" & Trim(!PartNum)
            .MoveNext
         Loop
         ClearResultSet RdoPcl
      End With
      If cmbPrt.ListCount > 0 Then bGoodRuns = GetRuns()
   Else
      MsgBox "No Matching Runs Recorded.", _
         vbInformation, Caption
   End If
   On Error Resume Next
   Set RdoPcl = Nothing
   cmbPrt.SetFocus
   Exit Sub
   
DiaErr1:
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub



Private Function GetRuns()
   'Dim RdoRns As ADODB.Recordset
   cmdSel.Enabled = True
   If sOldPart = cmbPrt Then Exit Function
   sOldPart = cmbPrt
   Grid1.Rows = 2
   Grid1.row = 1
   Grid1.Col = 0
   Grid1.Text = ""
   Grid1.Col = 1
   Grid1.Text = ""
   Grid1.Col = 2
   Grid1.Text = ""
   Grid1.Col = 3
   Grid1.Text = ""
   lblItems = 0
   cmbRun.Clear
   cmbPck.Clear
   Grid1.Enabled = False
   cmbPck.Enabled = False
   cmdCpl.Enabled = False
   optDel.Enabled = False
   lblDsc(1) = ""
   lblItm = ""
   lblQty = ""
   On Error GoTo DiaErr1
   sPartNumber = GetCurrentPart(cmbPrt, lblDsc(0))
   sSql = "SELECT RUNREF,RUNNO FROM RunsTable WHERE RUNREF='" _
          & Compress(cmbPrt) & "' AND (RUNSTATUS NOT LIKE 'C%')"
   LoadNumComboBox cmbRun, "####0", 1
   If bSqlRows Then
      cmbRun = cmbRun.List(0)
      GetRuns = True
   Else
      sPartNumber = ""
      cmdCpl.Enabled = False
      GetRuns = False
   End If
   GetStatus
   On Error Resume Next
   'Set RdoRns = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getruns"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub Grid1_Click()
   Dim iRow As Integer
   iRow = Grid1.row - 1
   cmbPck = vItems(iRow, 0)
   lblDsc(1) = vItems(iRow, 1)
   lblItm = vItems(iRow, 4)
   lblQty = Format(vItems(iRow, 5), ES_QuantityDataFormat)
   lblDate = vItems(iRow, 6)
   lblUom = vItems(iRow, 7)
   
End Sub

Private Sub Grid1_KeyPress(KeyAscii As Integer)
   Dim iRow As Integer
   If KeyAscii = 13 Then
      iRow = Grid1.row - 1
      cmbPck = vItems(iRow, 0)
      lblDsc(1) = vItems(iRow, 1)
      lblItm = vItems(iRow, 4)
      lblQty = Format(vItems(iRow, 5), ES_QuantityDataFormat)
      lblDate = vItems(iRow, 6)
      lblUom = vItems(iRow, 7)
   End If
   
End Sub


Private Sub lblDsc_Change(Index As Integer)
   If Left(lblDsc(0), 8) = "*** Part" Then
      lblDsc(0).ForeColor = ES_RED
   Else
      lblDsc(0).ForeColor = vbBlack
   End If
   
End Sub


'Add lots 4/29/02

Private Function CancelPick() As Byte
   Dim RdoPck As ADODB.Recordset
   Dim bLots As Byte
   Dim b As Byte
   Dim bByte As Byte
   
   Dim A As Integer
   Dim iRow As Integer
   Dim iTrans As Integer
   Dim iRef As Integer
   
   Dim lCOUNTER As Long
   Dim lLOTRECORD As Long
   Dim lSysCount As Long
   
   Dim cCost As Currency
   Dim cRemQty As Currency
   
   Dim clineCost As Currency
   Dim cLotQty As Currency
   Dim cOldCost As Currency
   Dim cQuantity As Currency
   Dim sMoPart As String * 31
   Dim sMoRun As String * 9
   Dim sPkPartRef As String
   Dim sPartNumber As String
   Dim sPkDate As String
   Dim sPkStatus As String
   
   On Error GoTo DiaErr1
   MouseCursor 13
   sPartNumber = Compress(cmbPrt)
   sSql = "SELECT PKPARTREF,PKMOPART,PKMORUN,PKMORUN," _
          & "PKADATE,PKAQTY,PKAMT FROM MopkTable WHERE (PKMOPART='" _
          & sPartNumber & "' AND PKMORUN=" & Val(cmbRun) & " AND " _
          & "PKRECORD=" & Val(lblItm) & ")"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPck, ES_FORWARD)
   If bSqlRows Then
      With RdoPck
         sMoPart = cmbPrt
         sPkDate = Format(!PKADATE, "mm/dd/yy")
         iRow = Len(Trim(str(cmbRun)))
         iRow = 5 - iRow
         sMoRun = "RUN" & Space$(iRow) & cmbRun
         iRow = 1
         vItems(iRow, 0) = "" & Trim(!PKPARTREF)
         sPartsGroup(iRow) = "" & Trim(!PKPARTREF)
         vItems(iRow, 1) = !PKAQTY
         vItems(iRow, 2) = !PKAMT
         cCost = cCost + !PKAMT
         ClearResultSet RdoPck
      End With
      If iRow > 0 Then
         A = iRow
         sPkPartRef = Compress(cmbPck)
         On Error Resume Next
         clsADOCon.ADOErrNum = 0
         clsADOCon.BeginTrans
         
         lCOUNTER = GetLastActivity()
         lSysCount = lCOUNTER + 1
         For iRow = 1 To A
            bByte = GetPartAccounts(sPartsGroup(iRow), sDebitAcct, sCreditAcct)
            cQuantity = Format(Val(vItems(iRow, 1)), ES_QuantityDataFormat)
            clineCost = Format((Val(vItems(iRow, 1)) * vItems(iRow, 2)), ES_QuantityDataFormat)
            If cQuantity > 0 Then
               'lCOUNTER = lCOUNTER + 1
               Dim strLoiNum As String
               Dim strPartRef As String
               Dim strLoiQty As String
               Dim strMOPartRef As String
               Dim strMORunNo As String
               Dim strLotUnitCost As String
               sPkPartRef = sPartsGroup(iRow)
               cOldCost = GetOldCost(sPkPartRef)
'               sSql = "INSERT INTO InvaTable (INTYPE,INPART,INREF1,INREF2,INPQTY,INAQTY,INAMT," _
'                      & "INCREDITACCT,INDEBITACCT,INMOPART,INMORUN,INNUMBER,INUSER)  " _
'                      & "VALUES(12,'" & sPkPartRef & "','CANCELED PICK','" & sMoPart & sMoRun & "'," _
'                      & cQuantity & "," & cQuantity & "," & cOldCost & ",'" _
'                      & sCreditAcct & "','" & sDebitAcct & "','" _
'                      & sPartNumber & "'," & Val(cmbRun) & "," & lCOUNTER & ",'" & sInitials & "')"
'               RdoCon.Execute sSql, rdExecDirect
               
               bLots = GetOpenLots(sPkPartRef, sPartNumber, Val(cmbRun), sPkDate)
               cLotQty = 0
               cRemQty = 0
               
               'insert lot transaction here
               For b = 1 To bLots
                  
                  If (Val(cLotQty) < Val(cQuantity)) Then
                     
                     lLOTRECORD = GetNextLotRecord(sLots(b, 0))
                     strLoiNum = CStr(sLots(b, 0))
                     strPartRef = CStr(sLots(b, 1))
                     strMOPartRef = CStr(sLots(b, 3))
                     strMORunNo = CStr(sLots(b, 4))
                         
                     cRemQty = Val(cQuantity) - Val(cLotQty)
                     If (Val(cRemQty) >= Val(sLots(b, 2))) Then
                        strLoiQty = CStr(sLots(b, 2))
                     Else
                        strLoiQty = CStr(Val(sLots(b, 2)) - Val(cRemQty))
                     End If
                     
                     cLotQty = Format(cLotQty + Val(strLoiQty), "######0.000")
                     
                     
                     lCOUNTER = lCOUNTER + 1
                     strLotUnitCost = GetLotUnitCost(strLoiNum, strPartRef, strMOPartRef, strMORunNo)
                     Dim totMatl As Currency
                     totMatl = GetTotMaterial(strPartRef)
                     sSql = "INSERT INTO InvaTable (INTYPE,INPART,INREF1,INREF2,INPQTY,INAQTY,INAMT,INTOTMATL," _
                            & "INCREDITACCT,INDEBITACCT,INLOTNUMBER,INMOPART,INMORUN,INNUMBER,INUSER)  " _
                            & "VALUES(12,'" & strPartRef & "','CANCELED PICK ITEM','" & sLots(b, 3) & Val(cmbRun) & "'," _
                            & Val(strLoiQty) & "," & Val(strLoiQty) & ",'" & Trim(strLotUnitCost) & "'," _
                            & CStr(totMatl) & ",'" _
                            & sCreditAcct & "','" & sDebitAcct & "','" & strLoiNum & "','" _
                            & strMOPartRef & "'," & Val(cmbRun) & "," & lCOUNTER & ",'" & sInitials & "')"
                     clsADOCon.ExecuteSql sSql
                      
                     
                     'sSql = "INSERT INTO LoitTable (LOINUMBER,LOIRECORD," _
                     '       & "LOITYPE,LOIPARTREF,LOIQUANTITY," _
                     '       & "LOIMOPARTREF,LOIMORUNNO," _
                     '       & "LOIACTIVITY,LOICOMMENT) " _
                     '       & "VALUES('" & strLoiNum & "'," _
                     '       & lLOTRECORD & ",12,'" & strPartRef & "'," _
                     '       & Val(strLoiQty) & ",'" & strMOPartRef & "'," & Val(cmbRun) & "," _
                     '       & lCOUNTER & ",'Canceled MO Pick')"
                     'clsADOCon.ExecuteSQL sSql
                     
                     sSql = "INSERT INTO LoitTable (LOINUMBER,LOIRECORD," _
                            & "LOITYPE,LOIPARTREF,LOIQUANTITY," _
                            & "LOIMOPARTREF,LOIMORUNNO," _
                            & "LOIACTIVITY,LOICOMMENT) " _
                            & "VALUES('" & strLoiNum & "'," _
                            & lLOTRECORD & ",12,'" & strPartRef & "'," _
                            & Val(strLoiQty) & ",'" & strMOPartRef & "'," & Val(cmbRun) & "," _
                            & lCOUNTER & ",'Canceled MO Pick Item')"
                     clsADOCon.ExecuteSql sSql 'rdExecDirect
                     
                     'Update the open lot in LoitTable talbe as MO canceled
                     sSql = "UPDATE LoitTable SET LOIMOPKCANCEL=1 " _
                              & " WHERE LOINUMBER='" & strLoiNum & "' AND " _
                              & " LOIPARTREF = '" & strPartRef & "' AND " _
                              & " LOIMOPARTREF = '" & strMOPartRef & "' AND " _
                              & " LOIMORUNNO = '" & strMORunNo & "' AND " _
                              & " LOIACTIVITY = '" & sLots(b, 5) & "'"
                              
                     clsADOCon.ExecuteSql sSql 'rdExecDirect
                     
                     'Update Lot Header
                     sSql = "UPDATE LohdTable SET LOTREMAININGQTY=LOTREMAININGQTY" _
                            & "+" & Val(strLoiQty) & " WHERE LOTNUMBER='" & strLoiNum & "'"
                     clsADOCon.ExecuteSql sSql
                  End If
               Next
               sSql = "UPDATE PartTable SET PAQOH=PAQOH+" & cQuantity _
                      & ",PALOTQTYREMAINING=PALOTQTYREMAINING+" & cQuantity _
                      & " WHERE PARTREF='" & sPartsGroup(iRow) & "' "
               clsADOCon.ExecuteSql sSql
            End If
            
            'Journal
            If iTrans > 0 And clineCost > 0 Then
               'Credit
               iRef = iRef + 1
               If Len(vItems(iRow, 0)) > 20 Then vItems(iRow, 0) = Left(vItems(iRow, 0), 20)
               sSql = "INSERT INTO JritTable (DCHEAD,DCTRAN,DCREF,DCCREDIT,DCACCTNO," _
                      & "DCDATE,DCDESCR,DCDESC,DCPARTNO,DCRUNNO) VALUES('" _
                      & sJournalID & "'," _
                      & iTrans & "," _
                      & iRef & "," _
                      & clineCost & ",'" _
                      & sCreditAcct & "','" _
                      & Format(ES_SYSDATE, "mm/dd/yy") & "','" _
                      & "CAPick" & "','" _
                      & vItems(iRow, 0) & "','" _
                      & sPartNumber & "'," _
                      & Val(cmbRun) & ")"
               clsADOCon.ExecuteSql sSql
               'Debit
               iRef = iRef + 1
               sSql = "INSERT INTO JritTable (DCHEAD,DCTRAN,DCREF,DCDEBIT,DCACCTNO," _
                      & "DCDATE,DCDESCR,DCDESC,DCPARTNO,DCRUNNO) VALUES('" _
                      & sJournalID & "'," _
                      & iTrans & "," _
                      & iRef & "," _
                      & clineCost & ",'" _
                      & sDebitAcct & "','" _
                      & Format(ES_SYSDATE, "mm/dd/yy") & "','" _
                      & "CAPick" & "','" _
                      & vItems(iRow, 0) & "','" _
                      & sPartNumber & "'," _
                      & Val(cmbRun) & ")"
               clsADOCon.ExecuteSql sSql
            End If
         Next
         sSql = "UPDATE RunsTable SET RUNSTATUS='PP'," _
                & "RUNCMATL=RUNCMATL-" & cCost & "," _
                & "RUNCOST=RUNCOST-" & cCost & " " _
                & "WHERE RUNREF='" & sPartNumber & "' " _
                & "AND RUNNO=" & Val(cmbRun) & " "
         clsADOCon.ExecuteSql sSql
         
         If optDel.Value = vbChecked Then
            sSql = "DELETE FROM MopkTable WHERE (PKMOPART='" & sPartNumber & "' " _
                   & "AND PKMORUN=" & Val(cmbRun) & " AND PKRECORD=" & Val(lblItm) & ")"
            clsADOCon.ExecuteSql sSql
         Else
            sSql = "UPDATE MopkTable SET PKTYPE=9,PKAQTY=0,PKAMT=0,PKADATE=NULL WHERE " _
                   & "(PKMOPART='" & sPartNumber & "' AND PKMORUN=" _
                   & Val(cmbRun) & " AND PKRECORD=" & Val(lblItm) & ")"
            clsADOCon.ExecuteSql sSql
         End If
         
         If clsADOCon.ADOErrNum = 0 Then
            clsADOCon.CommitTrans
            sSql = "UPDATE InvaTable SET INPDATE=INADATE WHERE INTYPE=12 AND " _
                   & "INPDATE IS NULL"
            clsADOCon.ExecuteSql sSql
            UpdateWipColumns lSysCount
            CancelPick = 1
         Else
            clsADOCon.RollbackTrans
            clsADOCon.ADOErrNum = 0
            CancelPick = 0
         End If
      Else
         On Error Resume Next
         sSql = "UPDATE RunsTable SET RUNSTATUS='PP'," _
                & "RUNCMATL=RUNCMATL-" & cCost & "," _
                & "RUNCOST=RUNCOST-" & cCost & " " _
                & "WHERE RUNREF='" & sPartNumber & "' " _
                & "AND RUNNO=" & Val(cmbRun) & " "
         clsADOCon.ExecuteSql sSql
         
         If optDel.Value = vbChecked Then
            sSql = "DELETE FROM MopkTable WHERE (PKMOPART='" & sPartNumber & "' " _
                   & "AND PKMORUN=" & Val(cmbRun) & " AND PKRECORD=" & Val(lblItm) & ")"
            clsADOCon.ExecuteSql sSql
         Else
            sSql = "UPDATE MopkTable SET PKTYPE=9,PKAQTY=0,PKAMT=0,PKADATE=NULL WHERE " _
                   & "(PKMOPART='" & sPartNumber & "' AND PKMORUN=" _
                   & Val(cmbRun) & " AND PKRECORD=" & Val(lblItm) & ")"
            clsADOCon.ExecuteSql sSql
         End If
         
         If clsADOCon.ADOErrNum = 0 Then
            'RdoCon.RollbackTrans
            clsADOCon.CommitTrans
            sSql = "UPDATE InvaTable SET INPDATE=INADATE WHERE " _
                   & "INTYPE=12 AND INPDATE IS NULL"
            clsADOCon.ExecuteSql sSql
            CancelPick = 1
         Else
            clsADOCon.RollbackTrans
            clsADOCon.ADOErrNum = 0
            CancelPick = 0
         End If
         MouseCursor 0
      End If
   Else
      sPkStatus = GetPickStatus()
      sSql = "UPDATE RunsTable SET RUNSTATUS='" & sPkStatus & "'," _
             & "RUNCMATL=RUNCMATL-" & cCost & "," _
             & "RUNCOST=RUNCOST-" & cCost & " " _
             & "WHERE RUNREF='" & sPartNumber & "' " _
             & "AND RUNNO=" & Val(cmbRun) & " "
      clsADOCon.ExecuteSql sSql
      MouseCursor 0
      CancelPick = 1
   End If
   Erase vItems
   cmdCpl.Enabled = False
   optDel.Enabled = False
   optDel.Value = vbUnchecked
   lblDsc(1) = ""
   lblItm = ""
   lblQty = ""
   Exit Function
   
DiaErr1:
   sProcName = "cancelpick"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function


Private Sub GetStatus()
   Dim RdoStu As ADODB.Recordset
   lblItems = 0
   cmdSel.Enabled = True
   On Error GoTo DiaErr1
   clsADOCon.ADOErrNum = 0
   sSql = "SELECT RUNREF,RUNSTATUS,RUNNO FROM RunsTable WHERE " _
          & "RUNREF = '" & Compress(cmbPrt) & "' AND RUNNO=" & Val(cmbRun) & " "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoStu, ES_FORWARD)
   If bSqlRows And clsADOCon.ADOErrNum = 0 Then
      lblStu = "" & Trim(RdoStu!RUNSTATUS)
   Else
      lblStu = ""
   End If
   Set RdoStu = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getstatus"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

'Lots 4/26/02 find lots for the Pick Transaction
'10/7/04 Added new Max() query

Private Function GetOpenLots(sLotPart As String, sMoNum As String, lMoRun As Long, _
                             sADate As Variant) As Byte
   Dim RdoPlot As ADODB.Recordset
   Dim b As Byte
   Dim iRow As Integer
   Dim iTotalLots As Integer
   Dim sOldLots(50, 2) As String
   Erase sLots
   GetOpenLots = 0
   On Error GoTo DiaErr1
   ' Removed LOIACTIVITY record - duplicate lot numbers.
'   sSql = "SELECT LOINUMBER, MAX(LOIRECORD) AS LOTRECORD" _
'          & " FROM LoitTable WHERE (LOIPARTREF='" _
'          & sLotPart & "' AND LOIMOPARTREF='" & sMoNum & "' AND " _
'          & "LOIMORUNNO=" & lMoRun & " AND LOITYPE=10) GROUP " _
'          & "BY LOINUMBER" ',LOIACTIVITY"

   sSql = "SELECT LOINUMBER, MAX(LOIRECORD) AS LOTRECORD" _
          & " FROM LoitTable WHERE (LOIPARTREF='" _
          & sLotPart & "' AND LOIMOPARTREF='" & sMoNum & "' AND " _
          & "LOIMORUNNO=" & lMoRun & " AND LOITYPE=10 AND " _
          & "(LOIMOPKCANCEL IS NULL OR LOIMOPKCANCEL <> 1)) GROUP " _
          & "BY LOINUMBER" ',LOIACTIVITY"
          
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPlot, ES_FORWARD)
   If bSqlRows Then
      With RdoPlot
         Do Until .EOF
            iTotalLots = iTotalLots + 1
            sOldLots(iTotalLots, 0) = "" & Trim(!LOINUMBER)
            sOldLots(iTotalLots, 1) = Trim$(str$(!LOTRECORD))
            .MoveNext
         Loop
         ClearResultSet RdoPlot
      End With
   End If
   For iRow = 1 To iTotalLots
'      sSql = "SELECT LOINUMBER,LOITYPE,LOIPARTREF,LOIADATE,LOIQUANTITY,LOIMOPARTREF," _
'             & "LOIMORUNNO FROM LoitTable WHERE (LOITYPE=10 AND LOINUMBER='" _
'             & sOldLots(iRow, 0) & "' AND LOIRECORD=" & sOldLots(iRow, 1) & ") "

      sSql = "SELECT LOINUMBER,LOITYPE,LOIPARTREF,LOIADATE,LOIQUANTITY,LOIMOPARTREF," _
             & "LOIMORUNNO,LOIACTIVITY FROM LoitTable WHERE (LOITYPE=10 AND LOINUMBER='" _
             & sOldLots(iRow, 0) & "' AND LOIRECORD=" & sOldLots(iRow, 1) & " AND " _
             & " (LOIMOPKCANCEL IS NULL OR LOIMOPKCANCEL <> 1))"
             
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoPlot, ES_FORWARD)
      If bSqlRows Then
         With RdoPlot
            Do Until .EOF
               GetOpenLots = GetOpenLots + 1
               sLots(GetOpenLots, 0) = "" & Trim(!LOINUMBER)
               sLots(GetOpenLots, 1) = "" & Trim(!LOIPARTREF)
               sLots(GetOpenLots, 2) = "" & Trim(str(Abs(!LOIQUANTITY)))
               sLots(GetOpenLots, 3) = "" & Trim(!LOIMOPARTREF)
               sLots(GetOpenLots, 4) = "" & Trim(str(!LOIMORUNNO))
               sLots(GetOpenLots, 5) = "" & Trim(str(!LoiActivity))
               .MoveNext
            Loop
            ClearResultSet RdoPlot
         End With
      End If
   Next
   Set RdoPlot = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getopenlots"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function


'0 = Part Number
'1 = Description
'2 = MoNumber
'3 = Run
'4 = Record
'5 = Quantity

Private Sub GetItems()
   Dim RdoItm As ADODB.Recordset
   Dim iItem As Integer
   Dim iRow As Integer
   cmbPck.Clear
   Grid1.Rows = 2
   Grid1.row = 1
   Grid1.Col = 0
   Grid1.Text = ""
   Grid1.Col = 1
   Grid1.Text = ""
   Grid1.Col = 2
   Grid1.Text = ""
   Grid1.Col = 3
   Grid1.Text = ""
   Erase vItems
   Erase sPartsGroup
   bGoodPick = 0
   On Error GoTo DiaErr1
   Grid1.Rows = 2
   iRow = 1
   sSql = "SELECT DISTINCT PKPARTREF,PKMOPART,PKMORUN,PKPDATE,PKRECORD," _
          & "PKAQTY,PKUNITS,PARTREF,PARTNUM,PADESC FROM MopkTable,PartTable WHERE " _
          & "(PKPARTREF=PARTREF AND PKMOPART='" & Compress(cmbPrt) & "' " _
          & "AND PKMORUN=" & Val(cmbRun) & " AND PKTYPE<12) ORDER BY " _
          & "PKPARTREF,PKRECORD"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoItm, ES_FORWARD)
   If bSqlRows Then
      With RdoItm
         Do Until .EOF
            If iItem > 0 Then iRow = iRow + 1
            vItems(iItem, 0) = "" & Trim(!PartNum)
            sPartsGroup(iItem) = "" & Trim(!PartRef)
            cmbPck.AddItem "" & Trim(!PartNum)
            vItems(iItem, 1) = "" & Trim(!PADESC)
            vItems(iItem, 2) = "" & Trim(!PKMOPART)
            vItems(iItem, 3) = Val(cmbRun)
            vItems(iItem, 4) = !PKRECORD
            vItems(iItem, 5) = Format(!PKAQTY, ES_QuantityDataFormat)
            vItems(iItem, 6) = Format(!PKPDATE, "mm/dd/yy")
            vItems(iItem, 7) = Trim(!PKUNITS)
            If iRow > 1 Then Grid1.Rows = Grid1.Rows + 1
            Grid1.row = iRow
            Grid1.Col = 0
            Grid1.Text = Format$(!PKRECORD)
            Grid1.Col = 1
            Grid1.Text = Trim(!PartNum)
            Grid1.Col = 2
            Grid1.Text = vItems(iItem, 6)
            Grid1.Col = 3
            Grid1.Text = vItems(iItem, 5) & " " & vItems(iItem, 7)
            If Not .EOF Then iItem = iItem + 1
            .MoveNext
         Loop
         ClearResultSet RdoItm
      End With
   End If
   lblItems = Grid1.Rows - 1
   If cmbPck.ListCount > 0 Then
      bGoodPick = 1
      Grid1.Enabled = True
      optDel.Enabled = True
      cmdCpl.Enabled = True
      cmbPck.Enabled = True
      cmbPck = cmbPck.List(0)
      lblDsc(1) = vItems(0, 1)
      lblItm = vItems(0, 4)
      lblQty = Format(vItems(0, 5), ES_QuantityDataFormat)
      lblDate = vItems(0, 6)
      lblUom = vItems(0, 7)
      cmdSel.Enabled = False
      cmbPck.SetFocus
   Else
      MsgBox "No Picked Items Were Found.", _
         vbInformation, Caption
   End If
   Set RdoItm = Nothing
   Exit Sub
DiaErr1:
   sProcName = "getitems"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

'10/25/03 Get the Run Status

Private Function GetPickStatus() As String
   Dim RdoPcs As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT COUNT(PKAQTY) FROM MopkTable WHERE PKAQTY=0 AND " _
          & "PKMOPART='" & Compress(cmbPrt) & "' AND PKMORUN=" & Val(cmbRun) & " "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPcs, ES_FORWARD)
   If bSqlRows Then
      If Not IsNull(RdoPcs.Fields(0)) Then
         If RdoPcs.Fields(0) = 0 Then
            GetPickStatus = "PC"
         Else
            sSql = "SELECT SUM(PKAQTY) FROM MopkTable WHERE PKMOPART='" _
                   & Compress(cmbPrt) & " AND PKMORUN=" & Val(cmbRun) & " "
            bSqlRows = clsADOCon.GetDataSet(sSql, RdoPcs, ES_FORWARD)
            If bSqlRows Then
               If Not IsNull(RdoPcs.Fields(0)) Then
                  If RdoPcs.Fields(0) > 0 Then GetPickStatus = "PP" _
                                       Else GetPickStatus = "PL"
               Else
                  GetPickStatus = "RL"
               End If
               ClearResultSet RdoPcs
            End If
         End If
      Else
         GetPickStatus = "RL"
      End If
   End If
   Set RdoPcs = Nothing
   Exit Function
   
DiaErr1:
   GetPickStatus = lblStu
   
End Function


'Gets Cost Recorded

Private Function GetOldCost(PartNumber As String) As Currency
   Dim RdoCost As ADODB.Recordset
   On Error Resume Next
   sSql = "SELECT INPART,INTYPE,INAMT,INMOPART,INMORUN FROM " _
          & "InvaTable WHERE (INPART='" & PartNumber & "' AND INTYPE=10 AND " _
          & "INMOPART='" & Compress(cmbPrt) & "' AND INMORUN=" _
          & Val(cmbRun) & ")"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCost, ES_FORWARD)
   If bSqlRows Then
      If Not IsNull(RdoCost!INAMT) Then GetOldCost = RdoCost!INAMT _
                    Else GetOldCost = 0
   End If
   Set RdoCost = Nothing
   
End Function

Private Function GetTotMaterial(PartNumber As String) As Currency
   Dim RdoCost As ADODB.Recordset
   On Error Resume Next
   GetTotMaterial = 0
   sSql = "SELECT INTOTMATL FROM " _
          & "InvaTable WHERE (INPART='" & PartNumber & "' AND INTYPE=10 AND " _
          & "INMOPART='" & Compress(cmbPrt) & "' AND INMORUN=" _
          & Val(cmbRun) & ")"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCost, ES_FORWARD)
   If bSqlRows Then
      GetTotMaterial = RdoCost!INTOTMATL
   End If
   Set RdoCost = Nothing
End Function


Private Function GetLotUnitCost(ByVal strLoiNum As String, ByVal strPartRef As String, _
                  ByVal strMOPartRef As String, _
                  strMORunNo As String) As Currency

   Dim RdoUnitCost As ADODB.Recordset
   On Error Resume Next
   
   sSql = "SELECT LOTUNITCOST FROM lohdTable WHERE " _
             & " LOTNUMBER = '" & strLoiNum & "'" _
            & " AND LOTPARTREF = '" & strPartRef & "'"
'            & " AND LOTMOPARTREF = '" & strMOPartRef & "'" _
'            & " AND LOTMORUNNO = '" & strMORunNo & "'"
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoUnitCost, ES_FORWARD)
   If bSqlRows Then
      If Not IsNull(RdoUnitCost!LotUnitCost) Then
         GetLotUnitCost = RdoUnitCost!LotUnitCost
      Else
         GetLotUnitCost = 0
      End If
   End If
   Set RdoUnitCost = Nothing

End Function

