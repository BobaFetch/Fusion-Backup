VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form DockODf01a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cancel an On Dock Delivery"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7305
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   7305
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdUp 
      DisabledPicture =   "DockODf01a.frx":0000
      DownPicture     =   "DockODf01a.frx":04F2
      Height          =   372
      Left            =   6400
      MaskColor       =   &H00000000&
      Picture         =   "DockODf01a.frx":09E4
      Style           =   1  'Graphical
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   4920
      Width           =   400
   End
   Begin VB.CommandButton cmdDn 
      DisabledPicture =   "DockODf01a.frx":0ED6
      DownPicture     =   "DockODf01a.frx":13C8
      Height          =   372
      Left            =   6840
      MaskColor       =   &H00000000&
      Picture         =   "DockODf01a.frx":18BA
      Style           =   1  'Graphical
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   4920
      Width           =   400
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "DockODf01a.frx":1DAC
      Style           =   1  'Graphical
      TabIndex        =   44
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   "C&ancel"
      Enabled         =   0   'False
      Height          =   315
      Left            =   5400
      TabIndex        =   21
      ToolTipText     =   "Cancel This Operation"
      Top             =   1200
      Width           =   875
   End
   Begin VB.ComboBox txtDte 
      Height          =   315
      Left            =   3120
      TabIndex        =   1
      Tag             =   "4"
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton cmdItm 
      Cancel          =   -1  'True
      Caption         =   "&Select"
      Height          =   315
      Left            =   5280
      TabIndex        =   2
      ToolTipText     =   "Fill With PO Items"
      Top             =   360
      Width           =   875
   End
   Begin VB.TextBox txtAcc 
      Enabled         =   0   'False
      Height          =   285
      Index           =   4
      Left            =   5280
      TabIndex        =   9
      Tag             =   "1"
      ToolTipText     =   "Quantity On The Vendor Pack Slip"
      Top             =   4080
      Width           =   1095
   End
   Begin VB.TextBox txtCmt 
      Enabled         =   0   'False
      Height          =   285
      Index           =   4
      Left            =   5280
      MultiLine       =   -1  'True
      TabIndex        =   10
      Tag             =   "3"
      ToolTipText     =   "Vendor Packing Slip Number"
      Top             =   4440
      Width           =   1935
   End
   Begin VB.TextBox txtAcc 
      Enabled         =   0   'False
      Height          =   285
      Index           =   3
      Left            =   5280
      TabIndex        =   7
      Tag             =   "1"
      ToolTipText     =   "Quantity On The Vendor Pack Slip"
      Top             =   3240
      Width           =   1095
   End
   Begin VB.TextBox txtCmt 
      Enabled         =   0   'False
      Height          =   285
      Index           =   3
      Left            =   5280
      MultiLine       =   -1  'True
      TabIndex        =   8
      Tag             =   "3"
      ToolTipText     =   "Vendor Packing Slip Number"
      Top             =   3600
      Width           =   1935
   End
   Begin VB.TextBox txtAcc 
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   5280
      TabIndex        =   5
      Tag             =   "1"
      ToolTipText     =   "Quantity On The Vendor Pack Slip"
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox txtCmt 
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   5280
      MultiLine       =   -1  'True
      TabIndex        =   6
      Tag             =   "3"
      ToolTipText     =   "Vendor Packing Slip Number"
      Top             =   2880
      Width           =   1935
   End
   Begin VB.TextBox txtAcc 
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   5280
      TabIndex        =   3
      Tag             =   "1"
      ToolTipText     =   "Quantity On The Vendor Pack Slip"
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox txtCmt 
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   5280
      MultiLine       =   -1  'True
      TabIndex        =   4
      Tag             =   "3"
      ToolTipText     =   "Vendor Packing Slip Number"
      Top             =   2160
      Width           =   1935
   End
   Begin VB.CommandButton cmdUpd 
      Caption         =   "&Apply"
      Enabled         =   0   'False
      Height          =   315
      Left            =   6360
      TabIndex        =   20
      ToolTipText     =   "Update Current List Of Items And Apply Changes"
      Top             =   1200
      Width           =   875
   End
   Begin VB.Frame z2 
      Height          =   60
      Left            =   240
      TabIndex        =   19
      Top             =   1080
      Width           =   7005
   End
   Begin VB.TextBox lblVendor 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   11
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
      Height          =   435
      Left            =   6360
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   1320
      Top             =   4800
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   5520
      FormDesignWidth =   7305
   End
   Begin VB.Image Dsdn 
      Height          =   300
      Left            =   240
      Picture         =   "DockODf01a.frx":255A
      Top             =   4800
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Endn 
      Height          =   300
      Left            =   480
      Picture         =   "DockODf01a.frx":2A4C
      Top             =   5160
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Enup 
      Height          =   300
      Left            =   240
      Picture         =   "DockODf01a.frx":2F3E
      Top             =   5160
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Dsup 
      Height          =   300
      Left            =   480
      Picture         =   "DockODf01a.frx":3430
      Top             =   4800
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Optional Function For Reporting Delivered Items Purposes Only"
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   5
      Left            =   1320
      TabIndex        =   43
      Top             =   0
      Width           =   5160
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      Height          =   285
      Index           =   7
      Left            =   2520
      TabIndex        =   42
      Top             =   400
      Width           =   1095
   End
   Begin VB.Label lblCmt 
      BackStyle       =   0  'Transparent
      Caption         =   "Pack Slip"
      Height          =   255
      Index           =   4
      Left            =   4320
      TabIndex        =   41
      Top             =   4440
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblCmt 
      BackStyle       =   0  'Transparent
      Caption         =   "Pack Slip"
      Height          =   255
      Index           =   3
      Left            =   4320
      TabIndex        =   40
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label lblCmt 
      BackStyle       =   0  'Transparent
      Caption         =   "Pack Slip"
      Height          =   255
      Index           =   2
      Left            =   4320
      TabIndex        =   39
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label lblCmt 
      BackStyle       =   0  'Transparent
      Caption         =   "Pack Slip"
      Height          =   255
      Index           =   1
      Left            =   4320
      TabIndex        =   38
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label lblPqt 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   4080
      TabIndex        =   37
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label lblPrt 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   960
      TabIndex        =   36
      ToolTipText     =   "Part Description"
      Top             =   4080
      Width           =   3105
   End
   Begin VB.Label lblItm 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   240
      TabIndex        =   35
      Top             =   4080
      Width           =   375
   End
   Begin VB.Label lblRev 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   610
      TabIndex        =   34
      Top             =   4080
      Width           =   285
   End
   Begin VB.Label lblRev 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   600
      TabIndex        =   33
      Top             =   3240
      Width           =   285
   End
   Begin VB.Label lblItm 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   240
      TabIndex        =   32
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label lblPrt 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   960
      TabIndex        =   31
      ToolTipText     =   "Part Description"
      Top             =   3240
      Width           =   3105
   End
   Begin VB.Label lblPqt 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   4080
      TabIndex        =   30
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label lblRev 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   600
      TabIndex        =   29
      Top             =   2520
      Width           =   285
   End
   Begin VB.Label lblItm 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   28
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label lblPrt 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   960
      TabIndex        =   27
      ToolTipText     =   "Part Description"
      Top             =   2520
      Width           =   3105
   End
   Begin VB.Label lblPqt 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   4080
      TabIndex        =   26
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Pack Slip Qty"
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
      Left            =   5280
      TabIndex        =   25
      Top             =   1560
      Width           =   1095
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
      Left            =   240
      TabIndex        =   24
      Top             =   1575
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
      Left            =   960
      TabIndex        =   23
      Top             =   1575
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
      Left            =   4080
      TabIndex        =   22
      Top             =   1575
      Width           =   1095
   End
   Begin VB.Label lblPqt 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   4095
      TabIndex        =   18
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label lblPrt 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   960
      TabIndex        =   17
      ToolTipText     =   "Part Description"
      Top             =   1800
      Width           =   3105
   End
   Begin VB.Label lblItm 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   16
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label lblRev 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   610
      TabIndex        =   15
      Top             =   1800
      Width           =   285
   End
   Begin VB.Label lblName 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3120
      TabIndex        =   14
      Top             =   720
      Width           =   3075
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "PO Number"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   13
      Top             =   400
      Width           =   1095
   End
End
Attribute VB_Name = "DockODf01a"
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

Dim iCurrPage As Integer
Dim iIndex As Integer
Dim iLastPage As Integer
Dim iTotalItems As Integer

Dim vItems(100, 10) As Variant
' 1 = PO
Private Const POITEM_PoNo = 1
' 2 = Item
Private Const POITEM_ItemNo = 2
' 3 = Item Rev
Private Const POITEM_ItemRev = 3
' 4 = Part Number
Private Const POITEM_PartNo = 4
' 5 = PO Qty
Private Const POITEM_OrderedQty = 5
' 6 = PS Qty
Private Const POITEM_ReceivedQty = 6
' 7 = PS Number
Private Const POITEM_PSNumber = 7
' 8 = ToolTipText
Private Const POITEM_ToolTip = 8
' 9 = Del Date
Private Const POITEM_DeliveredDate = 9

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd






Private Sub cmbPon_Click()
   GetCurrentVendor
   
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



Private Sub cmdDn_Click()
   'Next
   iCurrPage = iCurrPage + 1
   If iCurrPage > iLastPage Then iCurrPage = iLastPage
   GetTheNextGroup
   
End Sub

Private Sub cmdEnd_Click()
   Dim bResponse As Byte
   Dim sMsg As String
   
   sMsg = "Are You Sure That You Want To Cancel Without" & vbCr _
          & "Saving Any Changes To The Data?"
   bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
   If bResponse = vbYes Then
      ManageBoxes
      cmdUpd.Enabled = False
      cmdEnd.Enabled = False
      cmbPon.Enabled = True
      txtDte.Enabled = True
      cmdItm.Enabled = True
      On Error Resume Next
      cmbPon.SetFocus
   End If
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 5302
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub cmdItm_Click()
   GetPoItems
   
End Sub


Private Sub cmdUp_Click()
   'last
   iCurrPage = iCurrPage - 1
   If iCurrPage < 1 Then iCurrPage = 1
   GetTheNextGroup
   
End Sub

Private Sub cmdUpd_Click()
'   MsgBox "Requires Only That A Quantity Be Included " & vbCr _
'      & "And The Vendor Packing Slip Is Optional For " & vbCr _
'      & "Those Items To Be Reported As On Dock.", _
'      vbInformation, Caption
   UpdateList
   
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
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set DockODe02a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   lblVendor.BackColor = Es_FormBackColor
   txtDte = Format(ES_SYSDATE, "mm/dd/yy")
   For b = 1 To 4
      lblPrt(b).ToolTipText = "ToolTip = Part Description"
   Next
   cmdEnd.ToolTipText = "Cancel Work Not Updated And Return To Selection"
   
End Sub

Private Sub FillCombo()
   Dim RdoCmb As rdoResultset
   On Error GoTo DiaErr1
   sSql = "select DISTINCT PONUMBER from PoitTable" & vbCrLf _
          & "join PohdTable on PINUMBER = PONUMBER" & vbCrLf _
          & "where PITYPE = 14" & vbCrLf _
          & "order by PONUMBER desc "
   bSqlRows = GetDataSet(RdoCmb, ES_FORWARD)
   If bSqlRows Then
      With RdoCmb
         Do Until .EOF
            AddComboStr cmbPon.hWnd, "" & Format(!PONUMBER, "000000")
            .MoveNext
         Loop
         ClearResultSet RdoCmb
      End With
      If cmbPon.ListCount > 0 Then
         cmbPon = cmbPon.List(0)
         GetCurrentVendor
      End If
   End If
   Set RdoCmb = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
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


' 1 = PO
' 2 = Item
' 3 = Item Rev
' 4 = Part Number
' 5 = PO Qty
' 6 = PS Qty
' 7 = PS Number
' 8 = ToolTipText

Private Sub GetPoItems()
   Dim RdoGpi As rdoResultset
   Dim iRow As Integer
   Dim rPages As Single
   
   iIndex = -1
   iTotalItems = 0
   ManageBoxes
   iRow = 0
   On Error GoTo DiaErr1
   sSql = "SELECT PINUMBER,PIITEM,PIREV,PITYPE,PIPART,PIPQTY," _
      & "PIODDELPSNUMBER,PIODDELQTY,PIODDELDATE,PARTREF,PARTNUM,PADESC" & vbCrLf _
      & "FROM PoitTable" & vbCrLf _
      & "join PartTable on PIPART=PARTREF" & vbCrLf _
      & "where PITYPE=14 and PINUMBER=" & Val(cmbPon) & vbCrLf _
      & "and PIODDELIVERED=0" & vbCrLf _
      & "order by PIITEM, PIREV"
   bSqlRows = GetDataSet(RdoGpi, ES_FORWARD)
   If bSqlRows Then
      With RdoGpi
         Do Until .EOF
            iRow = iRow + 1
            iTotalItems = iTotalItems + 1
            vItems(iRow, POITEM_PoNo) = cmbPon
            vItems(iRow, POITEM_ItemNo) = !PIITEM
            vItems(iRow, POITEM_ItemRev) = Trim(!PIREV)
            vItems(iRow, POITEM_PartNo) = Trim(!PartNum)
            vItems(iRow, POITEM_OrderedQty) = Format(!PIPQTY, ES_QuantityDataFormat)
            vItems(iRow, POITEM_ReceivedQty) = Format(!PIODDELQTY, ES_QuantityDataFormat)
            vItems(iRow, POITEM_PSNumber) = "" & Trim(!PIODDELPSNUMBER)
            vItems(iRow, POITEM_ToolTip) = "" & Trim(!PADESC)
            If Not IsNull(!PIODDELDATE) Then
               vItems(iRow, POITEM_DeliveredDate) = "'" & Format(!PIODDELDATE, "mm/dd/yy") & "'"
            Else
               vItems(iRow, POITEM_DeliveredDate) = Format(Now, "mm/dd/yy")
            End If
            If iRow < 5 Then
               lblItm(iRow).Visible = True
               lblRev(iRow).Visible = True
               lblItm(iRow).Visible = True
               lblPrt(iRow).Visible = True
               lblPqt(iRow).Visible = True
               txtAcc(iRow).Visible = True
               txtCmt(iRow).Visible = True
               lblCmt(iRow).Visible = True
            End If
            .MoveNext
         Loop
         ClearResultSet RdoGpi
      End With
   End If
   
   If iTotalItems > 4 Then
      cmdUp.Enabled = True
      cmdUp.Picture = Enup.Picture
      cmdDn.Enabled = True
      cmdDn.Picture = Endn.Picture
   End If
   If iTotalItems > 0 Then
      iLastPage = 0.4 + (iTotalItems / 4)
      txtAcc(1).Enabled = True
      cmdItm.Enabled = False
      cmbPon.Enabled = False
      txtDte.Enabled = False
      cmdUpd.Enabled = True
      cmdEnd.Enabled = True
      iIndex = 0
      iCurrPage = 1
      For iRow = 1 To iTotalItems
         If iRow > 4 Then Exit For
         lblItm(iRow) = vItems(iRow, POITEM_ItemNo)
         lblRev(iRow) = vItems(iRow, POITEM_ItemRev)
         lblPrt(iRow) = vItems(iRow, POITEM_PartNo)
         lblPrt(iRow).ToolTipText = vItems(iRow, POITEM_ToolTip)
         lblPqt(iRow) = Format(vItems(iRow, POITEM_OrderedQty), ES_QuantityDataFormat)
         txtAcc(iRow) = Format(vItems(iRow, POITEM_ReceivedQty), ES_QuantityDataFormat)
         txtCmt(iRow) = vItems(iRow, POITEM_PSNumber)
         txtAcc(iRow).Enabled = True
         txtCmt(iRow).Enabled = True
      Next
      On Error Resume Next
      txtAcc(1).SetFocus
   End If
   Set RdoGpi = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getpoitems"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


' 1 = PO
' 2 = Item
' 3 = Item Rev
' 4 = Part Number
' 5 = PO Qty
' 6 = PS Qty
' 7 = PS Number
' 8 = ToolTipText

Private Sub GetTheNextGroup()
   Dim iList As Integer
   Dim iRow As Integer
   iIndex = (iCurrPage - 1) * 4
   ManageBoxes
   
   For iRow = iIndex + 1 To iTotalItems
      iList = iList + 1
      If iList > 4 Then Exit For
      lblItm(iList).Visible = True
      lblRev(iList).Visible = True
      lblItm(iList).Visible = True
      lblPrt(iList).Visible = True
      lblPqt(iList).Visible = True
      txtAcc(iList).Visible = True
      txtCmt(iList).Visible = True
      lblCmt(iList).Visible = True
      lblItm(iList) = vItems(iRow, POITEM_ItemNo)
      lblRev(iList) = vItems(iRow, POITEM_ItemRev)
      lblPrt(iList) = vItems(iRow, POITEM_PartNo)
      lblPrt(iList).ToolTipText = vItems(iRow, POITEM_ToolTip)
      lblPqt(iList) = Format(vItems(iRow, POITEM_OrderedQty), ES_QuantityDataFormat)
      txtAcc(iList) = Format(vItems(iRow, POITEM_ReceivedQty), ES_QuantityDataFormat)
      txtCmt(iList) = vItems(iRow, POITEM_PSNumber)
   Next
   
End Sub


Private Sub txtAcc_GotFocus(Index As Integer)
   SelectFormat Me
   
End Sub

Private Sub txtAcc_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub txtAcc_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyValue KeyAscii
   
End Sub

Private Sub txtAcc_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdUp_Click
   If KeyCode = vbKeyPageDown Then cmdDn_Click
   
End Sub

Private Sub txtAcc_Validate(Index As Integer, Cancel As Boolean)
   txtAcc(Index) = CheckLen(txtAcc(Index), 9)
   txtAcc(Index) = Format(Abs(Val(txtAcc(Index))), ES_QuantityDataFormat)
   vItems(Index + iIndex, POITEM_ReceivedQty) = Val(txtAcc(Index))
   
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


Private Sub txtCmt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdUp_Click
   If KeyCode = vbKeyPageDown Then cmdDn_Click
   
End Sub

Private Sub txtCmt_Validate(Index As Integer, Cancel As Boolean)
   txtCmt(Index) = CheckLen(txtCmt(Index), 20)
   vItems(iIndex + Index, POITEM_PSNumber) = txtCmt(Index)
   
End Sub

Private Sub txtDte_DropDown()
   ShowCalendar Me
   
End Sub


Private Sub txtDte_LostFocus()
   txtDte = CheckDate(txtDte)
   
End Sub












' 1 = PO
' 2 = Item
' 3 = Item Rev
' 4 = Part Number
' 5 = PO Qty
' 6 = PS Qty
' 7 = PS Number
' 8 = ToolTipText

Private Sub UpdateList()
   Dim iRow As Integer
   'Dim bResponse As Byte
   'Dim bDelivered As Byte
   'Dim cTotalCol As Currency
   Dim sMsg As String
   'Dim sPackSlip As String
   
   On Error GoTo DiaErr1
   
   'PO releases aren't currently used.  just assume release 0
   Dim PoRelease As Integer
   PoRelease = 0
   
   Dim PONUMBER As Long, POITEM As Integer, POREV As String
   PONUMBER = Val(cmbPon)
   
   'Confirm
   sMsg = "You Have Chosen To Update The PO Data With The Information" & vbCr _
          & "In The Selections.  Are You Sure That You Wish To Continue?"
   If MsgBox(sMsg, ES_YESQUESTION, Caption) <> vbYes Then
      CancelTrans
      Exit Sub
   End If
   
   'if partial receipts, confirm those
   Dim recQty As Currency, ordQty As Currency
   For iRow = 1 To iTotalItems
      ordQty = Val(vItems(iRow, POITEM_OrderedQty))
      recQty = Val(vItems(iRow, POITEM_ReceivedQty))
      POITEM = Val(vItems(iRow, POITEM_ItemNo))
      POREV = vItems(iRow, POITEM_ItemRev)
      
      If recQty > 0 Then
         If recQty < ordQty Then
            sMsg = "Only " & recQty & " of " & ordQty & " for item " & vItems(iRow, POITEM_ItemNo) & vItems(iRow, POITEM_ItemRev) _
               & " (" & vItems(iRow, POITEM_PartNo) & ") are being delivered.  " _
               & "Is this correct?"
            If MsgBox(sMsg, ES_YESQUESTION, Caption) <> vbYes Then
               Exit Sub
            End If
         
         ElseIf recQty > ordQty Then
            sMsg = recQty & "of item " & vItems(iRow, POITEM_ItemNo) & vItems(iRow, POITEM_ItemRev) _
               & " " & vItems(iRow, POITEM_PartNo) & " are being delivered.  " _
               & "This is more than than the ordered quantity of " & ordQty & ".  This is not allowed."
            MsgBox sMsg, vbExclamation, Caption
            Exit Sub
         End If
      End If
   Next
   
   MouseCursor 13
   RdoCon.BeginTrans
   For iRow = 1 To iTotalItems
      ordQty = Val(vItems(iRow, POITEM_OrderedQty))
      recQty = Val(vItems(iRow, POITEM_ReceivedQty))
      POITEM = Val(vItems(iRow, POITEM_ItemNo))
      POREV = vItems(iRow, POITEM_ItemRev)
      If recQty > 0 Then
'         sSql = "update PoitTable set PIODDELIVERED=1," & vbCrLf _
'            & "PIODDELDATE='" & vItems(iRow, POITEM_DeliveredDate) & "'," & vbCrLf _
'            & "PIODDELQTY=" & recQty & "," & vbCrLf _
'            & "PIODDELPSNUMBER='" & vItems(iRow, POITEM_PSNumber) & "'," & vbCrLf _
'            & "PIPQTY=" & recQty & "," & vbCrLf _
'            & "PIAQTY=" & recQty & "," & vbCrLf _
'            & "PIADATE='" & vItems(iRow, POITEM_DeliveredDate) & "'" & vbCrLf _
'            & "where PINUMBER=" & PoNumber & vbCrLf _
'            & "and PIRELEASE=" & PoRelease & vbCrLf _
'            & "and PIITEM=" & PoItem & vbCrLf _
'            & "and PIREV='" & PoRev & "'"
         sSql = "update PoitTable" & vbCrLf _
            & "set PIODDELIVERED=1," & vbCrLf _
            & "PIODDELDATE='" & vItems(iRow, POITEM_DeliveredDate) & "'," & vbCrLf _
            & "PIODDELQTY=" & recQty & "," & vbCrLf _
            & "PIODDELPSNUMBER='" & vItems(iRow, POITEM_PSNumber) & "'," & vbCrLf _
            & "PIPQTY=" & recQty & vbCrLf _
            & "where PINUMBER=" & PONUMBER & vbCrLf _
            & "and PIRELEASE=" & PoRelease & vbCrLf _
            & "and PIITEM=" & POITEM & vbCrLf _
            & "and PIREV='" & POREV & "'"
         RdoCon.Execute sSql, rdExecDirect
         
         'if this is a partial delivery, create a split PO item for the difference
         If recQty < ordQty Then
            Dim item As New ClassPoItem
            item.CreateSplit PONUMBER, PoRelease, POITEM, POREV, ordQty - recQty, False
         End If
      End If
   Next
   MouseCursor 0
   RdoCon.CommitTrans
   MsgBox "On dock items have been updated", vbInformation, Caption
   ManageBoxes
   cmdUpd.Enabled = False
   cmdEnd.Enabled = False
   cmbPon.Enabled = True
   txtDte.Enabled = True
   cmdItm.Enabled = True
   'On Error Resume Next
   cmbPon.SetFocus
   Exit Sub
   
DiaErr1:
   RdoCon.RollbackTrans
   sProcName = "updatelist"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub ManageBoxes()
   Dim iList As Integer
   For iList = 1 To 4
      lblItm(iList).Visible = False
      lblRev(iList).Visible = False
      lblPrt(iList).ToolTipText = "Part Description"
      lblPrt(iList).Visible = False
      lblPqt(iList).Visible = False
      txtAcc(iList).Visible = False
      txtCmt(iList).Visible = False
      lblCmt(iList).Visible = False
   Next
   
End Sub
