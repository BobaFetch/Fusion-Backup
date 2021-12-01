VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form LotsLTf04a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Return Inventory to Vendor"
   ClientHeight    =   4035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6990
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   6990
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   1320
      TabIndex        =   38
      Top             =   480
      Width           =   3015
   End
   Begin VB.CommandButton optPrn 
      Height          =   320
      Left            =   5520
      Picture         =   "LotsLTf04a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   37
      ToolTipText     =   "Print The Report"
      Top             =   3720
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   490
   End
   Begin VB.CommandButton optDis 
      Height          =   320
      Left            =   5040
      Picture         =   "LotsLTf04a.frx":018A
      Style           =   1  'Graphical
      TabIndex        =   36
      ToolTipText     =   "Display The Report"
      Top             =   3720
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   490
   End
   Begin VB.TextBox txtRMA 
      Height          =   315
      Left            =   1320
      TabIndex        =   4
      ToolTipText     =   "Enter RMA Number"
      Top             =   3480
      Width           =   3615
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   "C&ancel"
      Enabled         =   0   'False
      Height          =   315
      Left            =   6045
      TabIndex        =   7
      ToolTipText     =   "Cancel The Current Transacton"
      Top             =   480
      Visible         =   0   'False
      Width           =   875
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "LotsLTf04a.frx":0308
      Style           =   1  'Graphical
      TabIndex        =   27
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdFnd 
      DownPicture     =   "LotsLTf04a.frx":0AB6
      Height          =   315
      Left            =   4440
      Picture         =   "LotsLTf04a.frx":0DF8
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Find A Part Number"
      Top             =   480
      Visible         =   0   'False
      Width           =   350
   End
   Begin VB.TextBox cmbPrt1 
      Height          =   288
      Left            =   1320
      TabIndex        =   0
      Tag             =   "3"
      Top             =   480
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.ComboBox cmbVendor 
      Height          =   315
      Left            =   1320
      TabIndex        =   3
      Tag             =   "3"
      ToolTipText     =   "Select Or Enter Customer "
      Top             =   3000
      Width           =   1555
   End
   Begin VB.TextBox LotComment 
      Height          =   885
      Left            =   1800
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   23
      Text            =   "LotsLTf04a.frx":113A
      Top             =   6240
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.CommandButton cmdRtn 
      Caption         =   "R&eturn"
      Height          =   315
      Left            =   5925
      TabIndex        =   5
      ToolTipText     =   "Create The New Split Manufacturing Order"
      Top             =   3480
      Width           =   875
   End
   Begin VB.ComboBox cmbLot 
      Height          =   315
      Left            =   1320
      TabIndex        =   2
      Tag             =   "5"
      ToolTipText     =   "Select Lot From List And Press Select"
      Top             =   1320
      Width           =   3840
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   6045
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   7320
      Top             =   4200
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4035
      FormDesignWidth =   6990
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Release"
      Height          =   285
      Index           =   6
      Left            =   2880
      TabIndex        =   35
      Top             =   2520
      Width           =   585
   End
   Begin VB.Label lblPORel 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3525
      TabIndex        =   34
      ToolTipText     =   "Lot Location"
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label lblPORev 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   6285
      TabIndex        =   33
      ToolTipText     =   "Lot Location"
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Rev"
      Height          =   285
      Index           =   5
      Left            =   5880
      TabIndex        =   32
      Top             =   2520
      Width           =   345
   End
   Begin VB.Label lblPOItm 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   4920
      TabIndex        =   31
      ToolTipText     =   "Lot Location"
      Top             =   2520
      Width           =   615
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Item"
      Height          =   285
      Index           =   4
      Left            =   4440
      TabIndex        =   30
      Top             =   2520
      Width           =   465
   End
   Begin VB.Label lblPONum 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1320
      TabIndex        =   29
      ToolTipText     =   "Lot Location"
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "PO Number"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   28
      Top             =   2520
      Width           =   825
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "RMA Number"
      Height          =   285
      Index           =   8
      Left            =   240
      TabIndex        =   26
      Top             =   3480
      Width           =   1305
   End
   Begin VB.Label txtNme 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   3360
      TabIndex        =   25
      Top             =   3000
      Width           =   3375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Vendor"
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   24
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Location"
      Height          =   285
      Index           =   18
      Left            =   240
      TabIndex        =   22
      Top             =   2115
      Width           =   705
   End
   Begin VB.Label lblLotLoc 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1320
      TabIndex        =   21
      ToolTipText     =   "Lot Location"
      Top             =   2115
      Width           =   615
   End
   Begin VB.Label lblActCost 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.000"
      Height          =   285
      Left            =   5925
      TabIndex        =   20
      ToolTipText     =   "Costed Unit Value Of this Lot"
      Top             =   2115
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Actual Cost"
      Height          =   255
      Index           =   17
      Left            =   4800
      TabIndex        =   19
      Top             =   2115
      Width           =   855
   End
   Begin VB.Label lblStdCost 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.000"
      Height          =   285
      Left            =   5880
      TabIndex        =   18
      ToolTipText     =   "Part Number Standard Cost"
      Top             =   960
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Std Cost"
      Height          =   255
      Index           =   16
      Left            =   5040
      TabIndex        =   17
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "stuff down here V"
      Height          =   255
      Left            =   0
      TabIndex        =   16
      Top             =   5880
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblRem 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.000"
      Height          =   285
      Left            =   5925
      TabIndex        =   15
      ToolTipText     =   "Remaining In This Lot"
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Remaining"
      Height          =   255
      Index           =   3
      Left            =   4920
      TabIndex        =   14
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label lblLotSys 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1320
      TabIndex        =   13
      ToolTipText     =   "Existing System Lot Number"
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "System ID"
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   12
      Top             =   1680
      Width           =   1305
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Lots"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   11
      Top             =   1320
      Width           =   1305
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      Height          =   285
      Index           =   21
      Left            =   240
      TabIndex        =   10
      Top             =   480
      Width           =   1305
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1320
      TabIndex        =   9
      Top             =   960
      Width           =   3015
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   285
      Index           =   22
      Left            =   240
      TabIndex        =   8
      Top             =   960
      Width           =   1305
   End
End
Attribute VB_Name = "LotsLTf04a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007)) is the property of           ***
'*** ESI Software Engineering Inc, Stanwood, Washington, USA  ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
'8/23/05 New
'1/24/06 Made cmbPrt a TextBox (INTCOA Timeouts)
Option Explicit
Dim bCancel As Byte
Dim bOnLoad As Byte
Dim bUnLoad As Byte
Dim bView As Byte
Public bPrint As Byte


Dim cLotRemaining As Currency
Dim cStdCost As Currency
Dim cActCost As Currency
Dim cUnitCost As Currency
Dim cSplitQty As Currency

Dim sOldLot As String
Dim sOldPart As String

Dim sLots(100, 2) As String
Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd


Private Sub cmbVendor_Click()
   GetVendor
   
End Sub


Private Sub cmbVendor_LostFocus()
   GetVendor
   
End Sub
   

Private Sub cmbLot_LostFocus()
   Dim bByte As Byte
   Dim iList As Integer
   cmbLot = CheckLen(cmbLot, 40)
   For iList = 0 To cmbLot.ListCount - 1
      If cmbLot = cmbLot.List(iList) Then bByte = 1
   Next
   If bByte = 0 Then
      Beep
      If cmbLot.ListCount > 0 Then cmbLot = cmbLot.List(0)
   End If
   
   If cmbLot.ListCount > 0 And cmbLot.ListIndex < 0 _
         Then cmbLot.ListIndex = 0
   
   If cmbLot.ListCount > 0 Then
      lblLotSys = sLots(cmbLot.ListIndex, 0)
      GetThisLot
   End If
   
End Sub


Private Sub cmbPrt_Click()
   GetMONotPicked
End Sub


Private Sub cmbPrt_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF4 Then
      ViewParts.lblControl = "CMBPRT"
      ViewParts.txtPrt = cmbPrt
      ViewParts.Show
   End If
   
End Sub


Private Sub cmbPrt_LostFocus()
   cmbPrt = CheckLen(cmbPrt, 30)
   If bCancel = 1 Or bView = 1 Then Exit Sub
   
   If (Not ValidPartNumber(cmbPrt.Text)) Then
      MsgBox "Can't Select The Part Number Which Is Obsolete or Inactive. ", _
         vbInformation, Caption
      cmbPrt = ""
      Exit Sub
   End If
   
   If sOldPart <> cmbPrt Then GetMONotPicked
   
End Sub



Private Sub cmdCan_Click()
   Unload Me
   
End Sub



Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   bCancel = 1
   
End Sub


Private Sub cmdEnd_Click()
   Dim bResponse As Byte
   bResponse = MsgBox("Cancel The Creation Of The Lot Split.", _
               ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      cmbPrt.Enabled = True
      cmbLot.Enabled = True
      cmdEnd.Enabled = False
   End If
   
End Sub

Private Sub cmdFnd_Click()
   ViewParts.lblControl = "CMBPRT"
   ViewParts.txtPrt = cmbPrt
   'optVew.Value = vbChecked
   ViewParts.Show
   bView = 0
   
End Sub

Private Sub cmdFnd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   bView = 1
   
End Sub


Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext "5504"
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub cmdRtn_Click()
   ReturnLotPart
End Sub

Private Sub Form_Activate()
   Dim b As Byte
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      bUnLoad = 1
      cmdRtn.Enabled = False
      b = CheckLotStatus()
      FillCombo
      cmbPrt = ""
      lblDsc = ""
     If b = 1 Then
         FillVendor
         bOnLoad = 0
      Else
         MsgBox "Requires Lots Be Turned On.", _
            vbInformation, Caption
         Unload Me
      End If
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   If bUnLoad = 1 Then FormUnload
   Set LotsLTf04a = Nothing
   
End Sub


Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub


Private Sub FillCombo()
   On Error GoTo DiaErr1
   sSql = "SELECT PARTREF,PARTNUM FROM PartTable WHERE " _
          & "PALOTTRACK=1 AND PAINACTIVE = 0 AND PAOBSOLETE = 0 ORDER BY PARTREF"
   LoadComboBox cmbPrt
   If cmbPrt.ListCount > 0 Then cmbPrt = cmbPrt.List(0)
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

'1/24/05


Private Sub GetMONotPicked()
   Dim RdoPrt As ADODB.Recordset
   cmbLot.Clear
   sOldPart = cmbPrt
   On Error GoTo DiaErr1
   sSql = "SELECT PARTREF,PARTNUM,PADESC,PASTDCOST,PALOTQTYREMAINING " _
          & "FROM PartTable WHERE PARTREF='" & Compress(cmbPrt) & "' " _
          & "AND PALOTTRACK=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPrt, ES_FORWARD)
   If bSqlRows Then
      With RdoPrt
         cmbPrt = "" & Trim(!PartNum)
         lblDsc = "" & Trim(!PADESC)
         lblStdCost = Format(!PASTDCOST, "######0.000")
         lblRem = Format(!PALOTQTYREMAINING, "######0.000")
         ClearResultSet RdoPrt
      End With
   Else
      lblDsc = "*** Lot Tracked Part Number Not Found ***"
      lblRem = ""
      lblActCost = ""
      lblLotSys = ""
      lblLotLoc = ""
      lblPONum = ""
      lblPOItm = ""
      lblPORev = ""
      ' Now disable the return button
      cmdRtn.Enabled = False
   End If
   Set RdoPrt = Nothing
   If bSqlRows Then GetReceivedLots
   Exit Sub
   
DiaErr1:
   sProcName = "GetMONotPicked"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub Label2_Click()

End Sub

Private Sub lblActCost_Click()
   cActCost = Format(Val(lblActCost), ES_QuantityDataFormat)
   
End Sub

Private Sub lblDsc_Change()
   If Left(lblDsc, 6) = "*** Lo" Then _
           lblDsc.ForeColor = ES_RED Else _
           lblDsc.ForeColor = vbBlack
   
End Sub

Private Sub GetReceivedLots()
   Dim RdoLot As ADODB.Recordset
   Dim iRow As Integer
   On Error GoTo DiaErr1
   Erase lots
   iRow = -1
   
   'And PITYPE = 15
   sSql = "SELECT DISTINCT LOTPO, LOTPOITEM, LOTPOITEMREV, LOTNUMBER,LOTUSERLOTID," _
            & "LOTPARTREF,LOTREMAININGQTY FROM  LohdTable, poitTable a " _
      & " Where LOTPARTREF = '" & Compress(cmbPrt) & "' AND LOTREMAININGQTY > 0 " _
         & " AND LOTPO = PINUMBER AND LOTPOITEM = PIITEM AND LOTPOITEMREV = PIREV" _
      & " ORDER BY LOTPARTREF"
   
'   sSql = "SELECT LOTNUMBER,LOTUSERLOTID,LOTPARTREF,LOTREMAININGQTY FROM " _
'            & " LohdTable, InvaTable a " _
'            & " Where LOTPARTREF = '" & Compress(cmbPrt) & "' AND LOTREMAININGQTY > 0 " _
'            & " AND INTYPE = 15 AND a.INLOTNUMBER = LOTNUMBER "
'            & " AND INLOTNUMBER NOT IN ( " _
'               & "SELECT INLOTNUMBER FROM InvaTable b " _
'                  & " where a.INLOTNUMBER = b.INLOTNUMBER AND b.INTYPE = 10) " _
         & " ORDER BY a.inadate desc "
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoLot, ES_FORWARD)
   If bSqlRows Then
      With RdoLot
         Do Until .EOF
            AddComboStr cmbLot.hWnd, "" & Trim(!LOTUSERLOTID)
            iRow = iRow + 1
            sLots(iRow, 0) = "" & Trim(!lotNumber)
            sLots(iRow, 1) = "" & Trim(!LOTUSERLOTID)
            .MoveNext
         Loop
         ClearResultSet RdoLot
      End With
   End If
   
   If cmbLot.ListCount > 0 Then
      cmbLot = cmbLot.List(0)
      cmbLot.ListIndex = 0
      lblLotSys = sLots(0, 0)
      GetThisLot
   Else
      lblLotSys = "No Lots Found which could be returned."
   End If
   Set RdoLot = Nothing
   If bSqlRows Then GetThisLot
   Exit Sub
   
DiaErr1:
   sProcName = "GetReceivedLots"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub GetThisLot()
   Dim RdoLot As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT LOTNUMBER,LOTUSERLOTID,LOTPARTREF,LOTUNITCOST," _
          & "LOTLOCATION,LOTREMAININGQTY,LOTPO, LOTPOITEM, LOTPOITEMREV " _
          & " FROM LohdTable WHERE " _
          & "(LOTPARTREF='" & Compress(cmbPrt) & "' AND LOTNUMBER='" _
          & lblLotSys & "')"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoLot, ES_FORWARD)
   If bSqlRows Then
      With RdoLot
         lblRem = Format(!LOTREMAININGQTY, ES_QuantityDataFormat)
         lblActCost = Format(!LotUnitCost, ES_QuantityDataFormat)
         lblLotSys = "" & Trim(!lotNumber)
         lblLotLoc = "" & Trim(!LOTLOCATION)
         lblPONum = "" & Trim(!LOTPO)
         lblPOItm = "" & Trim(!LOTPOITEM)
         lblPORev = "" & Trim(!LOTPOITEMREV)
         Dim strVRef As String
         Dim strVName As String
         ' Get Customer full name
         GetPOVendor strVRef, strVName, lblPONum, lblPOItm
         If (strVRef <> "") Then
            cmbVendor = strVRef
            txtNme = strVName
         End If
         ' Now enable the return button
         cmdRtn.Enabled = True
         ClearResultSet RdoLot
      End With
   Else
      cmdRtn.Enabled = False
      lblLotSys = "No Lots With Quantities Found"
      lblRem = "0.000"
      lblActCost = ""
      lblLotSys = ""
      lblLotLoc = ""
      lblPONum = ""
      lblPOItm = ""
      lblPORev = ""
      ' Now disable the return button
      cmdRtn.Enabled = False
   End If
   Set RdoLot = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getthislot"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub
         
Private Sub GetPOVendor(ByRef strVRef As String, ByRef strVName As String, strPONum As String, strPOItm As String)
   Dim RdoPO As ADODB.Recordset
   
   sSql = "SELECT DISTINCT VEREF,VEBNAME FROM VndrTable, PohdTable WHERE PONUMBER ='" _
          & Compress(strPONum) & "' AND PohdTable.POVENDOR = VndrTable.VEREF"
          
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPO, ES_FORWARD)
   If bSqlRows Then
      strVRef = "" & Trim(RdoPO!VEREF)
      strVName = "" & Trim(RdoPO!VEBNAME)
      ClearResultSet RdoPO
   Else
      strVRef = ""
      strVName = ""
   End If

   Set RdoPO = Nothing
   
End Sub


Private Sub lblStdCost_Click()
   cStdCost = Format(Val(lblStdCost), ES_QuantityDataFormat)
End Sub

Public Sub ReturnLotPart()
   Dim bResponse As Byte
   Dim lCOUNTER As Long
   
   Dim lLOTRECORD As Long
   Dim lSysCount As Long
   
   Dim strLotNum As String
   Dim strPartNumber As String
   Dim strLotSys As String
   Dim strPONum As String
   Dim strRel As String
   Dim strPOItm As String
   Dim strPORev As String
   Dim strRem As String
   Dim vAdate As Variant
   Dim strVendor As String
   Dim strVdrName As String
   Dim strRMA As String
   
   
   strPartNumber = Compress(cmbPrt)
   strRem = "" & Trim(lblRem)
   strLotNum = Trim(cmbLot)
   strLotSys = "" & Trim(lblLotSys)
   strPONum = "" & Trim(lblPONum)
   'strRel = "" & Trim(lblPORel)
   strPOItm = "" & Trim(lblPOItm)
   strPORev = "" & Trim(lblPORev)
   strVendor = cmbVendor
   strVdrName = txtNme
   strRMA = txtRMA.Text
   
   vAdate = Format(GetServerDateTime(), "mm/dd/yy hh:mm")
   
   'Update the inavTable, Lothd and pohdTable
   On Error Resume Next
   
   Err.Clear
   clsADOCon.ADOErrNum = 0
   clsADOCon.BeginTrans
   
   ' Update the poit record to 40 (returned to customer)
   '" AND PIRELEASE=" & Val(strRel) & " "
   If (strPartNumber <> "" And strLotSys <> "" And strPONum <> "" And strRem <> "") Then
      
'      sSql = "UPDATE PoitTable SET PITYPE = 40 " _
'            & " WHERE PIPART = '" & strPartNumber & "' AND PINUMBER=" & Val(strPONum) _
'            & " AND (PIITEM='" & Val(strPOItm) & "' AND PIREV='" & strPORev & "')"
'
'      RdoCon.Execute sSql, rdExecDirect
      
      sSql = "UPDATE loitTable SET LOITYPE = 40, LOIQUANTITY = -1 * " & CStr(Val(strRem)) & "" _
            & " WHERE LOIPARTREF = '" & strPartNumber & "' AND LOINUMBER ='" & strLotSys & "'" _
            & " AND LOIPONUMBER = '" & Val(strPONum) & "' AND (LOIPOITEM='" & Val(strPOItm) & "' AND LOIPOREV='" & strPORev & "')"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "UPDATE lohdTable SET LOTREMAININGQTY = 0 " _
            & " WHERE LOTPARTREF = '" & strPartNumber & "' AND LOTNUMBER ='" & strLotSys & "'" _
            & " AND LOTPO = '" & Val(strPONum) & "' AND (LOTPOITEM ='" & Val(strPOItm) & "' AND LOTPOITEMREV ='" & strPORev & "')"
      clsADOCon.ExecuteSQL sSql
   
      sSql = "UPDATE InvaTable SET INTYPE = 40, INAQTY = -1 * " & CStr(Val(strRem)) & " , INREF1 = 'Part Returned' " _
            & " WHERE INPART = '" & strPartNumber & "' AND INLOTNUMBER ='" & strLotSys & "'" _
            & " AND INPONUMBER = '" & Val(strPONum) & "' AND (INPOITEM ='" & Val(strPOItm) & "' AND INPOREV ='" & strPORev & "')"
      clsADOCon.ExecuteSQL sSql
   
      sSql = "UPDATE PartTable SET PAQOH=PAQOH - " & Abs(strRem) & "," _
             & "PALOTQTYREMAINING=PALOTQTYREMAINING - " & Abs(strRem) & " " _
             & "WHERE PARTREF='" & strPartNumber & "'"
      clsADOCon.ExecuteSQL sSql
   End If

   
   If clsADOCon.ADOErrNum > 0 Then
      clsADOCon.RollbackTrans
      Exit Sub
   End If
   
   clsADOCon.CommitTrans
   
   ' Completed the update
   bResponse = MsgBox("Inventory return to Vendor is completed." _
            & vbCr _
            & "Do You Wish To print?", _
            ES_NOQUESTION, Caption)
   If bResponse = vbNo Then
      Exit Sub
   Else
      
      Load LotsLTf04b
      LotsLTf04b.lblPartNo = strPartNumber
      LotsLTf04b.lblLotNum = strLotNum
      LotsLTf04b.lblUserLotNum = strLotSys
      LotsLTf04b.lblLocation = lblLotLoc
      LotsLTf04b.lblQty = Trim(lblRem)
      LotsLTf04b.lblPONum = strPONum
      LotsLTf04b.lblPORel = ""
      LotsLTf04b.lblPOItm = lblPOItm
      LotsLTf04b.lblPORev = lblPORev
      LotsLTf04b.lblVendorName = strVdrName
      LotsLTf04b.lblRMA = strRMA
      
      Set LotsLTf04b.ParentForm = Me
      LotsLTf04b.Show vbModal
      If (bPrint = 1) Then
         PrintLabels strPartNumber, strLotSys, strVendor, strRMA, strRem
      End If
   End If
   
   
   
End Sub


Private Sub PrintLabels(strPartNumber As String, strLotSys As String, _
   strVendor As String, strRMA As String, strRem As String)
   
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   
   On Error GoTo DiaErr1
   MouseCursor 13
   
   optPrn.Value = True
   sCustomReport = GetCustomReport("invltf04")
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "RequestBy"
   aFormulaName.Add "Vendor"
   aFormulaName.Add "RMA"
   aFormulaName.Add "qty"
   
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'Requested By:" & CStr(sInitials) & "'")
   aFormulaValue.Add CStr("'" & CStr(strVendor) & "'")
   aFormulaValue.Add CStr("'" & CStr(strRMA) & "'")
   aFormulaValue.Add CStr("'" & strRem & "'")
    
    
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue

   sSql = "{lohdtable.LOTPARTREF} = '" & strPartNumber & "' AND {lohdtable.LOTNUMBER} = '" & strLotSys & "'"
         
   cCRViewer.SetReportSelectionFormula (sSql)
   cCRViewer.CRViewerSize Me
   cCRViewer.ShowGroupTree False
   cCRViewer.SetDbTableConnection

   cCRViewer.OpenCrystalReportObject Me, aFormulaName, 1, True
   
   cCRViewer.ClearFieldCollection aFormulaName
   cCRViewer.ClearFieldCollection aFormulaValue
   Set cCRViewer = Nothing
   
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "PrintLabels"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Private Sub FillVendor()
   On Error GoTo DiaErr1
   sSql = "SELECT DISTINCT VEREF,VENICKNAME FROM VndrTable ORDER BY VEREF"
          '& "CUALLOWTRANSFERS =1 ORDER BY CUREF"
   LoadComboBox cmbVendor
   If cmbVendor.ListCount = 0 Then
      MsgBox "There are no Vendors.", _
         vbInformation, Caption
   Else
      GetVendor
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "filltranscust"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub GetVendor()
   Dim rdoCst As ADODB.Recordset
   sSql = "SELECT VEBNAME FROM VndrTable WHERE VEREF='" _
          & Compress(cmbVendor) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoCst, ES_FORWARD)
   If bSqlRows Then txtNme = "" & Trim(rdoCst!VEBNAME) _
                             Else txtNme = ""
   Set rdoCst = Nothing
End Sub

Private Sub optDis_Click()
   PrintLabels "01602VAPZSJ", "40555-200201-42", "CROMPTON", "TEST1", 2
End Sub
