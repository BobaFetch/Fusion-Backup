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
   'Dim rdo As ADODB.Recordset
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
   clsADOCon.ExecuteSQL sSql
   If clsADOCon.RowsAffected > 0 Then
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
   Dim Ado As ADODB.Recordset
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
   bSqlRows = clsADOCon.GetDataSet(sSql, Ado, ES_FORWARD)
   If bSqlRows Then
      With Ado
         Do Until .EOF
            'AddComboStr cmbPon.hWnd, "" & Format(!PONUMBER, "000000")
            AddComboStr cmbPon.hwnd, "" & Format(!PINUMBER, "000000")
            .MoveNext
         Loop
         ClearResultSet Ado
      End With
      If cmbPon.ListCount > 0 Then
         cmbPon = cmbPon.List(0)
         GetCurrentVendor
      End If
      FillItemCombo
   End If
   Set Ado = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "FillPoCombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub



Private Sub FillItemCombo()
   Dim rdo As ADODB.Recordset
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
   bSqlRows = clsADOCon.GetDataSet(sSql, rdo, ES_FORWARD)
   If bSqlRows Then
      With rdo
         Do Until .EOF
            'AddComboStr cmbItem.hWnd, "" & Format(!PIITEM, "000") & !PIREV
            AddComboStr cmbItem.hwnd, "" & !PIITEM & !PIREV
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
   Dim RdoVnd As ADODB.Recordset
   sSql = "SELECT PONUMBER,POVENDOR,VEREF,VENICKNAME," _
          & "VEBNAME FROM PohdTable,VndrTable WHERE (VEREF=" _
          & "POVENDOR AND PONUMBER=" & Val(cmbPon) & ")"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoVnd, ES_FORWARD)
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


Private Sub txtCmt_GotFocus(Index As Integer)
   SelectFormat Me
End Sub

Private Sub txtCmt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub txtCmt_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyCase KeyAscii
   
End Sub



Private Sub GetItem()
   Dim rdo As ADODB.Recordset
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
   If clsADOCon.GetDataSet(sSql, rdo, ES_FORWARD) Then
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
   Set rdo = Nothing
End Sub
