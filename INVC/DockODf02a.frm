VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form DockODf02a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cancel an On Dock Delivery"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7680
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   7680
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "DockODf02a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   14
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox cmbItem 
      Height          =   315
      Left            =   3120
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Tag             =   "4"
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancelItem 
      Caption         =   "Cancel Item"
      Enabled         =   0   'False
      Height          =   450
      Left            =   6360
      TabIndex        =   9
      ToolTipText     =   "Update Current List Of Items And Apply Changes"
      Top             =   660
      Width           =   1200
   End
   Begin VB.TextBox lblVendor 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   720
      Width           =   1695
   End
   Begin VB.ComboBox cmbPon 
      Height          =   315
      Left            =   1320
      TabIndex        =   0
      Tag             =   "1"
      ToolTipText     =   "Contains Only PO's With On Dock Requirements"
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Close"
      Height          =   450
      Left            =   6360
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   1200
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   180
      Top             =   1020
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3030
      FormDesignWidth =   7680
   End
   Begin VB.Label lblDescription 
      BorderStyle     =   1  'Fixed Single
      Height          =   1125
      Left            =   2040
      TabIndex        =   16
      ToolTipText     =   "Part Description"
      Top             =   1800
      Width           =   5145
   End
   Begin VB.Label lblPart 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2040
      TabIndex        =   15
      ToolTipText     =   "Part Description"
      Top             =   1440
      Width           =   3105
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Item"
      Height          =   285
      Index           =   7
      Left            =   2520
      TabIndex        =   13
      Top             =   405
      Width           =   555
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Item/Rev"
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
      Left            =   1320
      TabIndex        =   12
      Top             =   1215
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number/Description                              "
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
      Index           =   1
      Left            =   2040
      TabIndex        =   11
      Top             =   1215
      Width           =   3135
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "PO Qty             "
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
      Left            =   5160
      TabIndex        =   10
      Top             =   1215
      Width           =   1095
   End
   Begin VB.Label lblQty 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5175
      TabIndex        =   8
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label lblItem 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1320
      TabIndex        =   7
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label lblRev 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1695
      TabIndex        =   6
      Top             =   1440
      Width           =   285
   End
   Begin VB.Label lblName 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3120
      TabIndex        =   5
      Top             =   720
      Width           =   3075
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "PO Number"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   400
      Width           =   1095
   End
End
Attribute VB_Name = "DockODf02a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
'10/28/04 renumbered lists
Option Explicit
Dim bCancel As Byte
Dim bOnLoad As Byte

'Dim iCurrPage As Integer
'Dim iIndex As Integer
'Dim iLastPage As Integer
'Dim iTotalItems As Integer

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd






Private Sub cmbItem_Click()
   GetItem
End Sub

Private Sub cmbItem_LostFocus()
   GetItem
End Sub

Private Sub cmbPon_Click()
   GetCurrentVendor
   FillItemCombo
End Sub


Private Sub cmbPon_LostFocus()
   Dim b As Byte
   Dim iList As Integer
   cmbPon = CheckLen(cmbPon, 6)
   cmbPon = Format(Abs(Val(cmbPon)), "000000")
   If bCancel Then Exit Sub
   For iList = 0 To cmbPon.ListCount - 1
      If cmbPon = cmbPon.List(iList) Then b = 1
   Next
   If b = 1 Then
      GetCurrentVendor
      FillItemCombo
   Else
      Beep
      MsgBox "The Requested PO Does Not Exist Or Is Not Listed", _
         vbInformation, Caption
      If cmbPon.ListCount > 0 Then cmbPon = cmbPon.List(0)
   End If
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub



Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   bCancel = 1
   
End Sub



'Private Sub cmdDn_Click()
'   'Next
'   iCurrPage = iCurrPage + 1
'   If iCurrPage > iLastPage Then iCurrPage = iLastPage
'   GetTheNextGroup
'
'End Sub

'Private Sub cmdEnd_Click()
'   Dim bResponse As Byte
'   Dim sMsg As String
'
'   sMsg = "Are You Sure That You Want To Cancel Without" & vbCr _
'          & "Saving Any Changes To The Data?"

Private Sub cmdCancelItem_Click()
   Select Case MsgBox("Cancel this on dock delivery?", vbQuestion + vbYesNo)
   Case vbYes
      CancelDelivery
   End Select
End Sub

Private Sub CancelDelivery()
   Dim rdo As rdoResultset
   Dim ItemNo As Integer, itemRev As String

   If cmbItem <> "" Then
      If IsNumeric(Right(cmbItem, 1)) Then
         ItemNo = cmbItem
      Else
         ItemNo = Left(cmbItem, Len(cmbItem) - 1)
         itemRev = Right(cmbItem, 1)
      End If
   End If
   
   sSql = "update PoitTable" & vbCrLf _
      & "set PIODDELIVERED = 0," & vbCrLf _
      & "PIODDELDATE = null," & vbCrLf _
      & "PIODDELQTY = 0," & vbCrLf _
      & "PIODDELPSNUMBER=''" & vbCrLf _
      & "where PINUMBER = " & Me.cmbPon & vbCrLf _
      & "and PIITEM = " & ItemNo & vbCrLf _
      & "and PIREV = '" & itemRev & "'" & vbCrLf
   RdoCon.Execute sSql
   If RdoCon.RowsAffected > 0 Then
      FillPoCombo
      MsgBox "On dock delivery canceled."
   Else
      MsgBox "Cancellation failed"
   End If
End Sub


'   bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
'   If bResponse = vbYes Then
'      ManageBoxes
'      cmdUpd.Enabled = False
'      cmdEnd.Enabled = False
'      cmbPon.Enabled = True
'      txtDte.Enabled = True
'      cmdItm.Enabled = True
'      On Error Resume Next
'      cmbPon.SetFocus
'   End If
'
'End Sub
'
Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 5302
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

'Private Sub cmdItm_Click()
'   GetPoItems
'
'End Sub


'Private Sub cmdUp_Click()
'   'last
'   iCurrPage = iCurrPage - 1
'   If iCurrPage < 1 Then iCurrPage = 1
'   GetTheNextGroup
'
'End Sub

'Private Sub cmdUpd_Click()
'   MsgBox "Requires Only That A Quantity Be Included " & vbCr _
'      & "And The Vendor Packing Slip Is Optional For " & vbCr _
'      & "Those Items To Be Reported As On Dock.", _
'      vbInformation, Caption
'   UpdateList
'
'End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillPoCombo
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   'FormatControls
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set DockODe02a = Nothing
   
End Sub



'Private Sub FormatControls()
'   Dim b As Byte
'   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
'   lblVendor.BackColor = Es_FormBackColor
'   txtDte = Format(ES_SYSDATE, "mm/dd/yy")
'   For b = 1 To 4
'      lblPrt(b).ToolTipText = "ToolTip = Part Description"
'   Next
'   cmdEnd.ToolTipText = "Cancel Work Not Updated And Return To Selection"
'
'End Sub
'
Private Sub FillPoCombo()
   Dim rdo As rdoResultset
   On Error GoTo DiaErr1
   
'   sSql = "update PoitTable" & vbCrLf _
'      & "set PIODDELIVERED=1," & vbCrLf _
'      & "PIODDELDATE='" & vItems(iRow, POITEM_DeliveredDate) & "'," & vbCrLf _
'      & "PIODDELQTY=" & recQty & "," & vbCrLf _
'      & "PIODDELPSNUMBER='" & vItems(iRow, POITEM_PSNumber) & "'," & vbCrLf _
'      & "PIPQTY=" & recQty & vbCrLf _
'      & "where PINUMBER=" & PONUMBER & vbCrLf _
'      & "and PIRELEASE=" & PoRelease & vbCrLf _
'      & "and PIITEM=" & POITEM & vbCrLf _
'      & "and PIREV='" & POREV & "'"
   
'   sSql = "select DISTINCT PONUMBER from PoitTable" & vbCrLf _
'          & "join PohdTable on PINUMBER = PONUMBER" & vbCrLf _
'          & "where PITYPE = 14" & vbCrLf _
'          & "order by PONUMBER desc "

   Me.cmbPon.Clear
   
   sSql = "select DISTINCT PINUMBER from PoitTable" & vbCrLf _
      & "where PITYPE = 14" & vbCrLf _
      & "and PIODDELIVERED = 1" & vbCrLf _
      & "and PIONDOCKINSPECTED = 0" & vbCrLf _
      & "order by PINUMBER desc"
   bSqlRows = GetDataSet(rdo, ES_FORWARD)
   If bSqlRows Then
      With rdo
         Do Until .EOF
            'AddComboStr cmbPon.hWnd, "" & Format(!PONUMBER, "000000")
            AddComboStr cmbPon.hWnd, "" & Format(!PINUMBER, "000000")
            .MoveNext
         Loop
         ClearResultSet rdo
      End With
      If cmbPon.ListCount > 0 Then
         cmbPon = cmbPon.List(0)
         GetCurrentVendor
      End If
      FillItemCombo
   End If
   Set rdo = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "FillPoCombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub



Private Sub FillItemCombo()
   Dim rdo As rdoResultset
   On Error GoTo DiaErr1
   
'   sSql = "select DISTINCT PONUMBER from PoitTable" & vbCrLf _
'          & "join PohdTable on PINUMBER = PONUMBER" & vbCrLf _
'          & "where PITYPE = 14" & vbCrLf _
'          & "order by PONUMBER desc "

   cmbItem.Clear
   If cmbPon = "" Then
      Exit Sub
   End If
   
   sSql = "select PIITEM, PIREV from PoitTable" & vbCrLf _
      & "where PINUMBER = " & Me.cmbPon & vbCrLf _
      & "and PITYPE = 14" & vbCrLf _
      & "and PIODDELIVERED = 1" & vbCrLf _
      & "and PIONDOCKINSPECTED = 0" & vbCrLf _
      & "order by PIITEM, PIREV"
   bSqlRows = GetDataSet(rdo, ES_FORWARD)
   If bSqlRows Then
      With rdo
         Do Until .EOF
            'AddComboStr cmbItem.hWnd, "" & Format(!PIITEM, "000") & !PIREV
            AddComboStr cmbItem.hWnd, "" & !PIITEM & !PIREV
            .MoveNext
         Loop
         ClearResultSet rdo
      End With
      If cmbItem.ListCount > 0 Then
         cmbItem = cmbItem.List(0)
         GetCurrentVendor
         'GetItem
      End If
   End If
   Set rdo = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "FillItemCombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub



Private Sub GetCurrentVendor()
   Dim RdoVnd As rdoResultset
   sSql = "SELECT PONUMBER,POVENDOR,VEREF,VENICKNAME," _
          & "VEBNAME FROM PohdTable,VndrTable WHERE (VEREF=" _
          & "POVENDOR AND PONUMBER=" & Val(cmbPon) & ")"
   bSqlRows = GetDataSet(RdoVnd, ES_FORWARD)
   If bSqlRows Then
      With RdoVnd
         lblVendor = "" & Trim(!VENICKNAME)
         lblName = "" & Trim(!VEBNAME)
         ClearResultSet RdoVnd
      End With
   Else
      lblVendor = ""
      lblName = "No Such PO Or PO Doesn't Qualify"
   End If
   Set RdoVnd = Nothing
   
End Sub


'' 1 = PO
'' 2 = Item
'' 3 = Item Rev
'' 4 = Part Number
'' 5 = PO Qty
'' 6 = PS Qty
'' 7 = PS Number
'' 8 = ToolTipText
'
'Private Sub GetPoItems()
'   Dim RdoGpi As rdoResultset
'   Dim iRow As Integer
'   Dim rPages As Single
'
'   iIndex = -1
'   iTotalItems = 0
'   ManageBoxes
'   iRow = 0
'   On Error GoTo DiaErr1
'   sSql = "SELECT PINUMBER,PIITEM,PIREV,PITYPE,PIPART,PIPQTY," _
'      & "PIODDELPSNUMBER,PIODDELQTY,PIODDELDATE,PARTREF,PARTNUM,PADESC" & vbCrLf _
'      & "FROM PoitTable" & vbCrLf _
'      & "join PartTable on PIPART=PARTREF" & vbCrLf _
'      & "where PITYPE=14 and PINUMBER=" & Val(cmbPon) & vbCrLf _
'      & "and PIODDELIVERED=0" & vbCrLf _
'      & "order by PIITEM, PIREV"
'   bSqlRows = GetDataSet(RdoGpi, ES_FORWARD)
'   If bSqlRows Then
'      With RdoGpi
'         Do Until .EOF
'            iRow = iRow + 1
'            iTotalItems = iTotalItems + 1
'            vItems(iRow, POITEM_PoNo) = cmbPon
'            vItems(iRow, POITEM_ItemNo) = !PIITEM
'            vItems(iRow, POITEM_ItemRev) = Trim(!PIREV)
'            vItems(iRow, POITEM_PartNo) = Trim(!PartNum)
'            vItems(iRow, POITEM_OrderedQty) = Format(!PIPQTY, ES_QuantityDataFormat)
'            vItems(iRow, POITEM_ReceivedQty) = Format(!PIODDELQTY, ES_QuantityDataFormat)
'            vItems(iRow, POITEM_PSNumber) = "" & Trim(!PIODDELPSNUMBER)
'            vItems(iRow, POITEM_ToolTip) = "" & Trim(!PADESC)
'            If Not IsNull(!PIODDELDATE) Then
'               vItems(iRow, POITEM_DeliveredDate) = "'" & Format(!PIODDELDATE, "mm/dd/yy") & "'"
'            Else
'               vItems(iRow, POITEM_DeliveredDate) = Format(Now, "mm/dd/yy")
'            End If
'            If iRow < 5 Then
'               lblItm(iRow).Visible = True
'               lblRev(iRow).Visible = True
'               lblItm(iRow).Visible = True
'               lblPrt(iRow).Visible = True
'               lblPqt(iRow).Visible = True
'               txtAcc(iRow).Visible = True
'               txtCmt(iRow).Visible = True
'               lblCmt(iRow).Visible = True
'            End If
'            .MoveNext
'         Loop
'         ClearResultSet RdoGpi
'      End With
'   End If
'
'   If iTotalItems > 4 Then
'      cmdUp.Enabled = True
'      cmdUp.Picture = Enup.Picture
'      cmdDn.Enabled = True
'      cmdDn.Picture = Endn.Picture
'   End If
'   If iTotalItems > 0 Then
'      iLastPage = 0.4 + (iTotalItems / 4)
'      txtAcc(1).Enabled = True
'      cmdItm.Enabled = False
'      cmbPon.Enabled = False
'      txtDte.Enabled = False
'      cmdUpd.Enabled = True
'      cmdEnd.Enabled = True
'      iIndex = 0
'      iCurrPage = 1
'      For iRow = 1 To iTotalItems
'         If iRow > 4 Then Exit For
'         lblItm(iRow) = vItems(iRow, POITEM_ItemNo)
'         lblRev(iRow) = vItems(iRow, POITEM_ItemRev)
'         lblPrt(iRow) = vItems(iRow, POITEM_PartNo)
'         lblPrt(iRow).ToolTipText = vItems(iRow, POITEM_ToolTip)
'         lblPqt(iRow) = Format(vItems(iRow, POITEM_OrderedQty), ES_QuantityDataFormat)
'         txtAcc(iRow) = Format(vItems(iRow, POITEM_ReceivedQty), ES_QuantityDataFormat)
'         txtCmt(iRow) = vItems(iRow, POITEM_PSNumber)
'         txtAcc(iRow).Enabled = True
'         txtCmt(iRow).Enabled = True
'      Next
'      On Error Resume Next
'      txtAcc(1).SetFocus
'   End If
'   Set RdoGpi = Nothing
'   Exit Sub
'
'DiaErr1:
'   sProcName = "getpoitems"
'   CurrError.Number = Err.Number
'   CurrError.Description = Err.Description
'   DoModuleErrors Me
'
'End Sub
'
'
'' 1 = PO
'' 2 = Item
'' 3 = Item Rev
'' 4 = Part Number
'' 5 = PO Qty
'' 6 = PS Qty
'' 7 = PS Number
'' 8 = ToolTipText
'
'Private Sub GetTheNextGroup()
'   Dim iList As Integer
'   Dim iRow As Integer
'   iIndex = (iCurrPage - 1) * 4
'   ManageBoxes
'
'   For iRow = iIndex + 1 To iTotalItems
'      iList = iList + 1
'      If iList > 4 Then Exit For
'      lblItm(iList).Visible = True
'      lblRev(iList).Visible = True
'      lblItm(iList).Visible = True
'      lblPrt(iList).Visible = True
'      lblPqt(iList).Visible = True
'      txtAcc(iList).Visible = True
'      txtCmt(iList).Visible = True
'      lblCmt(iList).Visible = True
'      lblItm(iList) = vItems(iRow, POITEM_ItemNo)
'      lblRev(iList) = vItems(iRow, POITEM_ItemRev)
'      lblPrt(iList) = vItems(iRow, POITEM_PartNo)
'      lblPrt(iList).ToolTipText = vItems(iRow, POITEM_ToolTip)
'      lblPqt(iList) = Format(vItems(iRow, POITEM_OrderedQty), ES_QuantityDataFormat)
'      txtAcc(iList) = Format(vItems(iRow, POITEM_ReceivedQty), ES_QuantityDataFormat)
'      txtCmt(iList) = vItems(iRow, POITEM_PSNumber)
'   Next
'
'End Sub
'

Private Sub txtAcc_GotFocus(Index As Integer)
   SelectFormat Me
   
End Sub

Private Sub txtAcc_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub txtAcc_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyValue KeyAscii
   
End Sub

'Private Sub txtAcc_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'   If KeyCode = vbKeyPageUp Then cmdUp_Click
'   If KeyCode = vbKeyPageDown Then cmdDn_Click
'
'End Sub

'Private Sub txtAcc_Validate(Index As Integer, Cancel As Boolean)
'   txtAcc(Index) = CheckLen(txtAcc(Index), 9)
'   txtAcc(Index) = Format(Abs(Val(txtAcc(Index))), ES_QuantityDataFormat)
'   vItems(Index + iIndex, POITEM_ReceivedQty) = Val(txtAcc(Index))
'
'End Sub


Private Sub txtCmt_GotFocus(Index As Integer)
   SelectFormat Me
End Sub

Private Sub txtCmt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub txtCmt_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyCase KeyAscii
   
End Sub


'Private Sub txtCmt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'   If KeyCode = vbKeyPageUp Then cmdUp_Click
'   If KeyCode = vbKeyPageDown Then cmdDn_Click
'
'End Sub
'
'Private Sub txtCmt_Validate(Index As Integer, Cancel As Boolean)
'   txtCmt(Index) = CheckLen(txtCmt(Index), 20)
'   vItems(iIndex + Index, POITEM_PSNumber) = txtCmt(Index)
'
'End Sub

'Private Sub txtDte_DropDown()
'   ShowCalendar Me
'
'End Sub
'
'
'Private Sub txtDte_LostFocus()
'   txtDte = CheckDate(txtDte)
'
'End Sub
'

' 1 = PO
' 2 = Item
' 3 = Item Rev
' 4 = Part Number
' 5 = PO Qty
' 6 = PS Qty
' 7 = PS Number
' 8 = ToolTipText

'Private Sub UpdateList()
'   Dim iRow As Integer
'   'Dim bResponse As Byte
'   'Dim bDelivered As Byte
'   'Dim cTotalCol As Currency
'   Dim sMsg As String
'   'Dim sPackSlip As String
'
'   On Error GoTo DiaErr1
'
'   'PO releases aren't currently used.  just assume release 0
'   Dim PoRelease As Integer
'   PoRelease = 0
'
'   Dim PONUMBER As Long, POITEM As Integer, POREV As String
'   PONUMBER = Val(cmbPon)
'
'   'Confirm
'   sMsg = "You Have Chosen To Update The PO Data With The Information" & vbCr _
'          & "In The Selections.  Are You Sure That You Wish To Continue?"
'   If MsgBox(sMsg, ES_YESQUESTION, Caption) <> vbYes Then
'      CancelTrans
'      Exit Sub
'   End If
'
'   'if partial receipts, confirm those
'   Dim recQty As Currency, ordQty As Currency
'   For iRow = 1 To iTotalItems
'      ordQty = Val(vItems(iRow, POITEM_OrderedQty))
'      recQty = Val(vItems(iRow, POITEM_ReceivedQty))
'      POITEM = Val(vItems(iRow, POITEM_ItemNo))
'      POREV = vItems(iRow, POITEM_ItemRev)
'
'      If recQty > 0 Then
'         If recQty < ordQty Then
'            sMsg = "Only " & recQty & " of " & ordQty & " for item " & vItems(iRow, POITEM_ItemNo) & vItems(iRow, POITEM_ItemRev) _
'               & " (" & vItems(iRow, POITEM_PartNo) & ") are being delivered.  " _
'               & "Is this correct?"
'            If MsgBox(sMsg, ES_YESQUESTION, Caption) <> vbYes Then
'               Exit Sub
'            End If
'
'         ElseIf recQty > ordQty Then
'            sMsg = recQty & "of item " & vItems(iRow, POITEM_ItemNo) & vItems(iRow, POITEM_ItemRev) _
'               & " " & vItems(iRow, POITEM_PartNo) & " are being delivered.  " _
'               & "This is more than than the ordered quantity of " & ordQty & ".  This is not allowed."
'            MsgBox sMsg, vbExclamation, Caption
'            Exit Sub
'         End If
'      End If
'   Next
'
'   MouseCursor 13
'   RdoCon.BeginTrans
'   For iRow = 1 To iTotalItems
'      ordQty = Val(vItems(iRow, POITEM_OrderedQty))
'      recQty = Val(vItems(iRow, POITEM_ReceivedQty))
'      POITEM = Val(vItems(iRow, POITEM_ItemNo))
'      POREV = vItems(iRow, POITEM_ItemRev)
'      If recQty > 0 Then
''         sSql = "update PoitTable set PIODDELIVERED=1," & vbCrLf _
''            & "PIODDELDATE='" & vItems(iRow, POITEM_DeliveredDate) & "'," & vbCrLf _
''            & "PIODDELQTY=" & recQty & "," & vbCrLf _
''            & "PIODDELPSNUMBER='" & vItems(iRow, POITEM_PSNumber) & "'," & vbCrLf _
''            & "PIPQTY=" & recQty & "," & vbCrLf _
''            & "PIAQTY=" & recQty & "," & vbCrLf _
''            & "PIADATE='" & vItems(iRow, POITEM_DeliveredDate) & "'" & vbCrLf _
''            & "where PINUMBER=" & PoNumber & vbCrLf _
''            & "and PIRELEASE=" & PoRelease & vbCrLf _
''            & "and PIITEM=" & PoItem & vbCrLf _
''            & "and PIREV='" & PoRev & "'"
'         sSql = "update PoitTable" & vbCrLf _
'            & "set PIODDELIVERED=1," & vbCrLf _
'            & "PIODDELDATE='" & vItems(iRow, POITEM_DeliveredDate) & "'," & vbCrLf _
'            & "PIODDELQTY=" & recQty & "," & vbCrLf _
'            & "PIODDELPSNUMBER='" & vItems(iRow, POITEM_PSNumber) & "'," & vbCrLf _
'            & "PIPQTY=" & recQty & vbCrLf _
'            & "where PINUMBER=" & PONUMBER & vbCrLf _
'            & "and PIRELEASE=" & PoRelease & vbCrLf _
'            & "and PIITEM=" & POITEM & vbCrLf _
'            & "and PIREV='" & POREV & "'"
'         RdoCon.Execute sSql, rdExecDirect
'
'         'if this is a partial delivery, create a split PO item for the difference
'         If recQty < ordQty Then
'            Dim item As New ClassPoItem
'            item.CreateSplit PONUMBER, PoRelease, POITEM, POREV, ordQty - recQty, False
'         End If
'      End If
'   Next
'   MouseCursor 0
'   RdoCon.CommitTrans
'   MsgBox "On dock items have been updated", vbInformation, Caption
'   ManageBoxes
'   cmdUpd.Enabled = False
'   cmdEnd.Enabled = False
'   cmbPon.Enabled = True
'   txtDte.Enabled = True
'   cmdItm.Enabled = True
'   'On Error Resume Next
'   cmbPon.SetFocus
'   Exit Sub
'
'DiaErr1:
'   RdoCon.RollbackTrans
'   sProcName = "updatelist"
'   CurrError.Number = Err.Number
'   CurrError.Description = Err.Description
'   DoModuleErrors Me
'
'End Sub
'
'Private Sub ManageBoxes()
'   Dim iList As Integer
'   For iList = 1 To 4
'      lblItm(iList).Visible = False
'      lblRev(iList).Visible = False
'      lblPrt(iList).ToolTipText = "Part Description"
'      lblPrt(iList).Visible = False
'      lblPqt(iList).Visible = False
'      txtAcc(iList).Visible = False
'      txtCmt(iList).Visible = False
'      lblCmt(iList).Visible = False
'   Next
'
'End Sub

Private Sub GetItem()
   Dim rdo As rdoResultset
   Dim ItemNo As Integer, itemRev As String
'   ItemNo = Left(cmbItem, 3)
'   If Len(cmbItem) > 3 Then
'      itemRev = Right(cmbItem, 1)
'   Else
'      itemRev = ""
'   End If

   If cmbItem <> "" Then
      If IsNumeric(Right(cmbItem, 1)) Then
         ItemNo = cmbItem
      Else
         ItemNo = Left(cmbItem, Len(cmbItem) - 1)
         itemRev = Right(cmbItem, 1)
      End If
   End If
   
   sSql = "select RTRIM(PARTNUM) as PARTNUM, RTRIM(PADESC) as PADESC," & vbCrLf _
      & "PIODDELQTY" & vbCrLf _
      & "from PoitTable" & vbCrLf _
      & "join PartTable on PIPART = PARTREF" & vbCrLf _
      & "where PINUMBER = " & Me.cmbPon & vbCrLf _
      & "and PIITEM = " & ItemNo & vbCrLf _
      & "and PIREV = '" & itemRev & "'" & vbCrLf
   If GetDataSet(rdo, ES_FORWARD) Then
      lblItem = ItemNo
      lblRev = itemRev
      lblPart = rdo!PartNum
      lblDescription = rdo!PADESC
      lblQty = rdo!PIODDELQTY
      cmdCancelItem.Enabled = True
   Else
      lblItem = ""
      lblRev = ""
      lblPart = ""
      lblDescription = ""
      lblQty = ""
      cmdCancelItem.Enabled = False
   End If
End Sub
