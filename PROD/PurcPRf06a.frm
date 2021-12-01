VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form PurcPRf06a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Revise A PO Line Item Price (Invoiced Item)"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6465
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   6465
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "PurcPRf06a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   20
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CheckBox optStatus 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   255
      Left            =   4920
      TabIndex        =   19
      Top             =   1800
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtQty 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      ToolTipText     =   "Actual Quantity"
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox txtVnd 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   4
      ToolTipText     =   "Vendor"
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton cmdUpd 
      Caption         =   "&Apply"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5520
      TabIndex        =   5
      ToolTipText     =   "Update And Apply Changes"
      Top             =   600
      Width           =   855
   End
   Begin VB.TextBox txtPrc 
      Height          =   285
      Left            =   5280
      TabIndex        =   3
      Tag             =   "1"
      ToolTipText     =   "New Price"
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox txtRev 
      Height          =   285
      Left            =   4200
      TabIndex        =   2
      Tag             =   "3"
      ToolTipText     =   "Line Item Revision (If Any)"
      Top             =   1080
      Width           =   375
   End
   Begin VB.TextBox txtItm 
      Height          =   285
      Left            =   2880
      TabIndex        =   1
      Tag             =   "1"
      ToolTipText     =   "Line Item"
      Top             =   1080
      Width           =   495
   End
   Begin VB.TextBox txtPon 
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Tag             =   "1"
      ToolTipText     =   "Purchase Order"
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5520
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
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
      FormDesignHeight=   3120
      FormDesignWidth =   6465
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Actual)"
      Height          =   255
      Index           =   7
      Left            =   2520
      TabIndex        =   18
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label lblPart 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1320
      TabIndex        =   17
      ToolTipText     =   "Line Item Part Number"
      Top             =   2520
      Width           =   3255
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   15
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity"
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   14
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Enter All Areas"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   600
      Width           =   3855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Price"
      Height          =   255
      Index           =   4
      Left            =   4680
      TabIndex        =   12
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label lblVnd 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1320
      TabIndex        =   11
      ToolTipText     =   "Vendor"
      Top             =   1800
      Width           =   3255
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Vendor"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   10
      Tag             =   "VEmd"
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Rev"
      Height          =   255
      Index           =   2
      Left            =   3480
      TabIndex        =   9
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Item"
      Height          =   255
      Index           =   1
      Left            =   2400
      TabIndex        =   8
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "PO Number"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   7
      Top             =   1080
      Width           =   1095
   End
End
Attribute VB_Name = "PurcPRf06a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007)) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
'5/20/04 New
'1/11/06 Removed Help jump
Option Explicit
Dim bOnLoad As Byte
Dim bGoodPo As Byte
Dim bStatus As Byte

Dim lOldPo As Long

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmdCan_Click()
   Unload Me
   
End Sub



Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 4355
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub cmdUpd_Click()
   Dim bResponse As Byte
   Dim sMsg As String
   
   If bStatus <> 17 Then
      MsgBox "This Item Has Not Been Invoiced.", _
         vbInformation, Caption
   Else
      sMsg = "Are You Certain That You Want To Update " & vbCr _
             & "The Planned Cost (Price) For This Item?"
      bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
      If bResponse = vbYes Then
         On Error Resume Next
         clsADOCon.ADOErrNum = 0
         sSql = "UPDATE PoitTable SET PIESTUNIT=" & Val(txtPrc) _
                & " WHERE (PINUMBER=" & Val(txtPon) & " AND " _
                & "PIITEM=" & Val(txtItm) & " AND PIREV='" _
                & Trim(txtRev) & "')"
         clsADOCon.ExecuteSQL sSql
         cmdUpd.Enabled = False
         If clsADOCon.ADOErrNum = 0 Then
            SysMsg "The Item Price Was Updated.", True
         Else
            MsgBox "Could Not Update The Item Price.", _
               vbInformation, Caption
         End If
      Else
         CancelTrans
      End If
   End If
   
End Sub

Private Sub Form_Activate()
   Dim b As Byte
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      b = GetPermission()
      ES_PurchasedDataFormat = GetPODataFormat()
   End If
   MouseCursor 0
   If b = 0 Then
      MsgBox "Permissions Have Not Been Set For This Function. Search " & vbCr _
         & "Help For " & Caption & ".", _
         vbInformation, Caption
      Unload Me
   End If
   bOnLoad = 0
   
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
   Set PurcPRf06a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtPrc = "0.000"
   txtVnd.BackColor = Me.BackColor
   lOldPo = 9999999
   
End Sub




Private Sub lblPart_Change()
   If Left(lblPart, 8) = "*** Line" Then _
           lblPart.ForeColor = ES_RED Else lblPart.ForeColor = vbBlack
   
End Sub

Private Sub lblVnd_Change()
   If Left(lblVnd, 9) = "*** Purch" Then _
           lblVnd.ForeColor = ES_RED Else lblVnd.ForeColor = vbBlack
   
End Sub

Private Sub txtItm_LostFocus()
   txtItm = CheckLen(txtItm, 3)
   txtItm = Format(Abs(Val(txtItm)), "##0")
   
End Sub


Private Sub txtPon_LostFocus()
   txtPon = CheckLen(txtPon, 6)
   txtPon = Format(Abs(Val(txtPon)), "000000")
   bGoodPo = GetPo()
   
End Sub


Private Sub txtPrc_LostFocus()
   txtPrc = CheckLen(txtPrc, 9)
   txtPrc = Format(Abs(Val(txtPrc)), ES_PurchasedDataFormat)
   
End Sub



Private Function GetPo() As Byte
   Dim RdoPon As ADODB.Recordset
   On Error GoTo DiaErr1
   If Val(txtPon) > 0 And Val(txtPon) <> lOldPo Then
      sSql = "SELECT PONUMBER,POVENDOR,VEREF,VENICKNAME,VEBNAME " _
             & "FROM PohdTable,VndrTable WHERE (PONUMBER=" _
             & Val(txtPon) & " AND POVENDOR=VEREF)"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoPon, ES_FORWARD)
      If bSqlRows Then
         With RdoPon
            txtVnd = "" & Trim(!VENICKNAME)
            lblVnd = "" & Trim(!VEBNAME)
            ClearResultSet RdoPon
            GetPo = 1
         End With
      Else
         GetPo = 0
         txtVnd = ""
         lblVnd = "*** Purchase Order Wasn't Found ***"
      End If
   End If
   cmdUpd.Enabled = False
   optStatus.Visible = False
   bStatus = 0
   txtItm = ""
   txtRev = ""
   txtQty = ""
   txtPrc = ""
   lblPart = ""
   If Val(txtPon) > 0 Then lOldPo = Val(txtPon)
   Set RdoPon = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "GetPo    "
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub txtRev_LostFocus()
   txtRev = CheckLen(txtRev, 2)
   If lblVnd.ForeColor <> ES_RED Then bStatus = GetLineItem()
   
End Sub



Private Function GetLineItem() As Byte
   Dim RdoItm As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT PINUMBER,PIITEM,PIREV,PITYPE,PIPART," _
          & "PIAQTY,PIESTUNIT,PARTREF,PARTNUM FROM " _
          & "PoitTable,PartTable WHERE " _
          & "(PIPART=PARTREF AND PINUMBER=" & Val(txtPon) & " " _
          & "AND PIITEM=" & Val(txtItm) & " AND PIREV='" _
          & Trim(txtRev) & "' AND PITYPE<>16)"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoItm, ES_FORWARD)
   If bSqlRows Then
      With RdoItm
         GetLineItem = !PITYPE
         txtPrc = Format(!PIESTUNIT, ES_PurchasedDataFormat)
         txtQty = Format(!PIAQTY, "######0.000")
         lblPart = "" & Trim(!PIPART)
         ClearResultSet RdoItm
         optStatus.Visible = True
         optStatus.Caption = ""
         cmdUpd.Enabled = True
      End With
   Else
      GetLineItem = 0
      txtQty = ""
      txtPrc = ""
      lblPart = "*** Line Item Not Found Or Is Canceled ***"
   End If
   Select Case GetLineItem
      Case 14
         optStatus.Caption = "Open Item"
      Case 15
         optStatus.Caption = "Received"
      Case 17
         optStatus.Caption = "Received"
      Case 18
         optStatus.Caption = "On Dock"
      Case Else
         optStatus.Visible = False
   End Select
   Set RdoItm = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getlineitem"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Function GetPermission() As Byte
   Dim RdoPon As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT AllowPostInvoicePricing FROM Preferences WHERE PreRecord=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPon, ES_FORWARD)
   If bSqlRows Then
      With RdoPon
         If Not IsNull(!AllowPostInvoicePricing) Then _
                       GetPermission = !AllowPostInvoicePricing Else _
                       GetPermission = 0
         ClearResultSet RdoPon
      End With
   End If
   Set RdoPon = Nothing
   Exit Function
   
DiaErr1:
   GetPermission = 0
End Function
