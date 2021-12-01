VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form jevRTe01c 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Inspection Tag Items"
   ClientHeight    =   5325
   ClientLeft      =   2100
   ClientTop       =   1635
   ClientWidth     =   7140
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   Icon            =   "jevRTe01c.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5325
   ScaleWidth      =   7140
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdLst 
      Caption         =   "<< &Last"
      Enabled         =   0   'False
      Height          =   300
      Left            =   4680
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   120
      Width           =   732
   End
   Begin VB.CommandButton cmdNxt 
      Caption         =   "& Next >>"
      Enabled         =   0   'False
      Height          =   300
      Left            =   5440
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   120
      Width           =   732
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "jevRTe01c.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   34
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox cmbCas 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1560
      TabIndex        =   8
      Tag             =   "8"
      ToolTipText     =   "Select Disposition Code From List"
      Top             =   4440
      Width           =   1675
   End
   Begin VB.ComboBox cmbRes 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1320
      TabIndex        =   11
      Tag             =   "8"
      ToolTipText     =   "Select Resposibility Code From List"
      Top             =   5400
      Visible         =   0   'False
      Width           =   1675
   End
   Begin VB.ComboBox cmbDis 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   4440
      TabIndex        =   4
      Tag             =   "8"
      ToolTipText     =   "Select Disposition Code From List"
      Top             =   840
      Width           =   1675
   End
   Begin VB.ComboBox txtDte 
      Height          =   315
      Left            =   6000
      TabIndex        =   9
      Tag             =   "4"
      Top             =   4440
      Width           =   1095
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6960
      Top             =   5040
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   5325
      FormDesignWidth =   7140
   End
   Begin VB.TextBox txtScr 
      Height          =   285
      Left            =   4920
      TabIndex        =   2
      Tag             =   "1"
      Top             =   480
      Width           =   915
   End
   Begin VB.TextBox txtRwk 
      Height          =   285
      Left            =   3120
      TabIndex        =   1
      Tag             =   "1"
      Top             =   480
      Width           =   915
   End
   Begin VB.TextBox txtQty 
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Tag             =   "1"
      Top             =   480
      Width           =   915
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "&Delete"
      Height          =   285
      Left            =   6240
      TabIndex        =   24
      TabStop         =   0   'False
      ToolTipText     =   "Delete The Current Item"
      Top             =   920
      Width           =   875
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   285
      Left            =   6240
      TabIndex        =   20
      TabStop         =   0   'False
      ToolTipText     =   "Add An Item"
      Top             =   600
      Width           =   875
   End
   Begin VB.TextBox txtCor 
      Height          =   825
      Left            =   3600
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Tag             =   "9"
      Top             =   3510
      Width           =   3435
   End
   Begin VB.TextBox txtTst 
      Height          =   825
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Tag             =   "9"
      Top             =   3510
      Width           =   3315
   End
   Begin VB.TextBox txtDip 
      Height          =   1395
      Left            =   3600
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Tag             =   "9"
      Top             =   1770
      Width           =   3435
   End
   Begin VB.TextBox txtDis 
      Height          =   1395
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Tag             =   "9"
      Top             =   1770
      Width           =   3315
   End
   Begin VB.ComboBox cmbCde 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1200
      TabIndex        =   3
      Tag             =   "8"
      ToolTipText     =   "Select Characteristic Code From List"
      Top             =   840
      Width           =   1675
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   6240
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   90
      Width           =   875
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cause Code"
      Height          =   285
      Index           =   12
      Left            =   120
      TabIndex        =   33
      Top             =   4440
      Width           =   1545
   End
   Begin VB.Label lblCause 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   120
      TabIndex        =   32
      Top             =   4800
      Width           =   3060
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Responsibility"
      Height          =   285
      Index           =   11
      Left            =   0
      TabIndex        =   31
      Top             =   5400
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Label lblRes 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3120
      TabIndex        =   30
      Top             =   5400
      Visible         =   0   'False
      Width           =   2820
   End
   Begin VB.Label lblDis 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3120
      TabIndex        =   29
      Top             =   1200
      Width           =   3060
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Disposition Code"
      Height          =   285
      Index           =   10
      Left            =   3120
      TabIndex        =   28
      Top             =   840
      Width           =   1545
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Scrap"
      Height          =   285
      Index           =   9
      Left            =   4200
      TabIndex        =   27
      Top             =   480
      Width           =   705
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Rework"
      Height          =   285
      Index           =   8
      Left            =   2400
      TabIndex        =   26
      Top             =   480
      Width           =   765
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity"
      Height          =   285
      Index           =   7
      Left            =   120
      TabIndex        =   25
      Top             =   480
      Width           =   975
   End
   Begin VB.Label lblTag 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   285
      Left            =   1200
      TabIndex        =   23
      Top             =   120
      Width           =   1365
   End
   Begin VB.Label lblItem 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1"
      Height          =   285
      Left            =   3120
      TabIndex        =   22
      Top             =   120
      Width           =   375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Item"
      Height          =   285
      Index           =   6
      Left            =   2640
      TabIndex        =   21
      Top             =   120
      Width           =   615
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Corrective Action Date"
      Height          =   285
      Index           =   5
      Left            =   4080
      TabIndex        =   19
      Top             =   4440
      Width           =   1830
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Corrective Action"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   4
      Left            =   3600
      TabIndex        =   18
      Top             =   3240
      Width           =   2445
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cause "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   3
      Left            =   120
      TabIndex        =   17
      Top             =   3240
      Width           =   2445
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Disposition Instructions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   2
      Left            =   3600
      TabIndex        =   16
      Top             =   1560
      Width           =   2445
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description of Discrepancy"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   15
      Top             =   1560
      Width           =   2445
   End
   Begin VB.Label lblCde 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   120
      TabIndex        =   14
      Top             =   1200
      Width           =   2820
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Discrepancy"
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   13
      Top             =   840
      Width           =   1185
   End
End
Attribute VB_Name = "jevRTe01c"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
Option Explicit
'//JEVCO Custom 10/14/02
Dim AdoItems As ADODB.Recordset

Dim bOnLoad As Byte
Dim iItemIndex As Integer
Dim iItemCount As Integer
Dim iLastItem As Integer

Dim cQuantity As Currency
Dim sCompTag As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub cmbCas_Click()
   GetCause
End Sub


Private Sub cmbCas_LostFocus()
   sSql = "UPDATE RjitTable SET RITCAUSE='" & Compress(cmbCas) & "' WHERE " _
          & "RITREF='" & sCompTag & "' AND RITITM=" & str(lblItem)
   clsADOCon.ExecuteSQL sSql
   
End Sub


Private Sub cmbCde_Click()
   If Not bOnLoad Then GetCharaCode
   
End Sub


Private Sub cmbCde_LostFocus()
   cmbCde = CheckLen(cmbCde, 12)
   GetCharaCode
   On Error Resume Next
   sSql = "UPDATE RjitTable SET RITCHARCODE='" & Compress(cmbCde) & "' WHERE " _
          & "RITREF='" & sCompTag & "' AND RITITM=" & str(lblItem)
   clsADOCon.ExecuteSQL sSql
   
   
End Sub


Private Sub cmbDis_Click()
   GetDispCode
   
End Sub


Private Sub cmbDis_LostFocus()
   cmbDis = CheckLen(cmbDis, 12)
   GetDispCode
   On Error Resume Next
   sSql = "UPDATE RjitTable SET RITDISPCODE='" & Compress(cmbDis) & "' WHERE " _
          & "RITREF='" & sCompTag & "' AND RITITM=" & str(lblItem)
   clsADOCon.ExecuteSQL sSql
   
End Sub


Private Sub cmbRes_Click()
   GetRespCode
   
End Sub


Private Sub cmbRes_LostFocus()
   cmbRes = CheckLen(cmbRes, 12)
   GetRespCode
   On Error Resume Next
   sSql = "UPDATE RjitTable SET RITRESPCODE='" & Compress(cmbRes) & "' WHERE " _
          & "RITREF='" & sCompTag & "' AND RITITM=" & str(lblItem)
   clsADOCon.ExecuteSQL sSql
   
End Sub


Private Sub cmdAdd_Click()
   Additems
   
End Sub

Private Sub cmdCan_Click()
   Unload Me
   
End Sub


Private Sub cmdDel_Click()
   DeleteItem
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 6102
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub

Private Sub cmdLst_Click()
   If iItemIndex > 1 Then
      iItemIndex = iItemIndex - 1
      cmdNxt.Enabled = True
   Else
      iItemIndex = 1
      cmdLst.Enabled = False
      cmdNxt.Enabled = True
   End If
   GetCurrentItem
   
End Sub

Private Sub cmdNxt_Click()
   If iItemIndex < iItemCount Then
      iItemIndex = iItemIndex + 1
      cmdLst.Enabled = True
   Else
      iItemIndex = iItemCount
      cmdNxt.Enabled = False
      cmdLst.Enabled = True
   End If
   GetCurrentItem
   
End Sub

Private Sub Form_Activate()
   MouseCursor 0
   If bOnLoad Then
      FillCombo
      If iItemCount = 0 Then GetItems
      bOnLoad = 0
   End If
   
End Sub

Private Sub Form_Load()
   SetFormSize Me
   FormatControls
   Move 200, 600
   iItemCount = 0
   bOnLoad = 1
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   InspRTe01b.optItm.value = vbUnchecked
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   MdiSect.Enabled = True
   Set AdoItems = Nothing
   Set jevRTe01c = Nothing
   
End Sub



Private Sub GetItems()
   sCompTag = Compress(lblTag)
   MouseCursor 13
   On Error GoTo DiaErr1
   sSql = "SELECT RITREF,RITITM,RITCHARCODE FROM RjitTable WHERE RITREF='" _
          & sCompTag & "' ORDER BY RITITM"
   bSqlRows = clsADOCon.GetDataSet(sSql, AdoItems, ES_KEYSET)
   If bSqlRows Then
      With AdoItems
         lblItem = Format(!RITITM, "#0")
         cmbCde = "" & Trim(!RITCHARCODE)
         Do Until AdoItems.EOF
            iItemCount = iItemCount + 1
            iLastItem = !RITITM
            .MoveNext
         Loop
         ClearResultSet AdoItems
      End With
      GetCharaCode
      If iItemCount > 1 Then cmdNxt.Enabled = True
   Else
      AdoItems.AddNew
      AdoItems!RITREF = sCompTag
      AdoItems!RITITM = 1
      AdoItems.Update
      bSqlRows = clsADOCon.GetDataSet(sSql, AdoItems, ES_KEYSET)
      If bSqlRows Then iLastItem = 1
   End If
   iItemIndex = 1
   GetCurrentItem
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "getitems"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub FillCombo()
   On Error GoTo DiaErr1
   sSql = "Qry_FillDescripancyCodes"
   LoadComboBox cmbCde
   
   sSql = "Qry_FillDispositionCodes"
   LoadComboBox cmbDis
   If ES_CUSTOM = "JEVCO" Then
      sSql = "SELECT CAUSEREF,CAUSENUM FROM RjcaTable "
      LoadComboBox cmbCas
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Private Sub Additems()
   Dim bResponse As Byte
   
   If bResponse = vbYes Then
      MouseCursor 13
      On Error GoTo DiaErr1
      iLastItem = iLastItem + 1
      iItemCount = iItemCount + 1
      AdoItems.AddNew
      AdoItems!RITREF = sCompTag
      AdoItems!RITITM = iLastItem
      AdoItems!RITDATE = Null
      AdoItems.Update
      
      lblItem = str(iLastItem)
      cmbCde = ""
      txtDis = ""
      txtDip = ""
      txtCor = ""
      txtTst = ""
      txtDte = ""
      iItemIndex = iLastItem
      cmdLst.Enabled = True
      cmdNxt.Enabled = False
   Else
      CancelTrans
   End If
   On Error Resume Next
   MouseCursor 0
   cmbCde.SetFocus
   Exit Sub
   
DiaErr1:
   CurrError.Description = Err.Description
   MouseCursor 0
   MsgBox CurrError.Description & vbCr & " Can't Add Record.", vbExclamation, Caption
   
End Sub

Private Sub lblCde_Change()
   If Left(lblCde, 8) = "*** Char" Then
      lblCde.ForeColor = ES_RED
   Else
      lblCde.ForeColor = vbBlack
   End If
   
End Sub

Private Sub lblDis_Change()
   If Left(lblDis, 8) = "*** Disp" Then
      lblDis.ForeColor = ES_RED
   Else
      lblDis.ForeColor = vbBlack
   End If
   
End Sub

Private Sub lblRes_Change()
   If Left(lblDis, 8) = "*** Resp" Then
      lblRes.ForeColor = ES_RED
   Else
      lblRes.ForeColor = vbBlack
   End If
   
End Sub

Private Sub txtCor_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdLst_Click
   If KeyCode = vbKeyPageDown Then cmdNxt_Click
   
End Sub


Private Sub txtCor_LostFocus()
   On Error Resume Next
   txtCor = CheckLen(txtCor, 1020)
   txtCor = StrCase(txtCor, ES_FIRSTWORD)
   sSql = "UPDATE RjitTable SET RITCORA='" & txtCor & "' WHERE " _
          & "RITREF='" & sCompTag & "' AND RITITM=" & str(lblItem)
   clsADOCon.ExecuteSQL sSql
   
End Sub


Private Sub txtDip_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdLst_Click
   If KeyCode = vbKeyPageDown Then cmdNxt_Click
   
End Sub


Private Sub txtDip_LostFocus()
   On Error Resume Next
   txtDip = CheckLen(txtDip, 1020)
   txtDip = StrCase(txtDip, ES_FIRSTWORD)
   sSql = "UPDATE RjitTable SET RITDISP='" & txtDip & "' WHERE " _
          & "RITREF='" & sCompTag & "' AND RITITM=" & str(lblItem)
   clsADOCon.ExecuteSQL sSql
   
End Sub


Private Sub txtDis_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdLst_Click
   If KeyCode = vbKeyPageDown Then cmdNxt_Click
   
End Sub


Private Sub txtDis_LostFocus()
   On Error Resume Next
   txtDis = CheckLen(txtDis, 1020)
   txtDis = StrCase(txtDis, ES_FIRSTWORD)
   sSql = "UPDATE RjitTable SET RITDESC='" & txtDis & "' WHERE " _
          & "RITREF='" & sCompTag & "' AND RITITM=" & str(lblItem)
   clsADOCon.ExecuteSQL sSql
   
   
End Sub


Private Sub txtDte_DropDown()
   ShowCalendar Me, 2500
   
End Sub

Private Sub txtDte_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdLst_Click
   If KeyCode = vbKeyPageDown Then cmdNxt_Click
   
End Sub

Private Sub txtDte_LostFocus()
   On Error Resume Next
   If Len(Trim(txtDte)) > 0 Then
      If Trim(cmbDis) = "" Then
         MsgBox "Requires A Disposition Code.", _
            vbExclamation, Caption
         txtDte = ""
         Exit Sub
      End If
      txtDte = CheckDate(txtDte)
      sSql = "UPDATE RjitTable SET RITDATE='" & txtDte _
             & "',RITACT=1 WHERE " _
             & "RITREF='" & sCompTag & "' AND RITITM=" & str(lblItem)
   Else
      txtDte = ""
      sSql = "UPDATE RjitTable SET RITDATE=Null" _
             & ",RITACT=0 WHERE " _
             & "RITREF='" & sCompTag & "' AND RITITM=" & str(lblItem)
   End If
   clsADOCon.ExecuteSQL sSql
   
End Sub


Private Sub txtQty_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdLst_Click
   If KeyCode = vbKeyPageDown Then cmdNxt_Click
   
End Sub

Private Sub txtQty_LostFocus()
   txtQty = CheckLen(txtQty, 10)
   txtQty = Format(txtQty, ES_QuantityDataFormat)
   cQuantity = Val(txtQty)
   On Error Resume Next
   sSql = "UPDATE RjitTable SET RITQTY=" & txtQty & " WHERE " _
          & "RITREF='" & sCompTag & "' AND RITITM=" & str(lblItem)
   clsADOCon.ExecuteSQL sSql
   
End Sub


Private Sub txtRwk_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdLst_Click
   If KeyCode = vbKeyPageDown Then cmdNxt_Click
   
End Sub

Private Sub txtRwk_LostFocus()
   txtRwk = CheckLen(txtRwk, 10)
   txtRwk = Format(txtRwk, ES_QuantityDataFormat)
   If Val(txtRwk) > cQuantity Then
      txtRwk = Format(cQuantity, ES_QuantityDataFormat)
      txtScr = Format(0, ES_QuantityDataFormat)
   End If
   On Error Resume Next
   sSql = "UPDATE RjitTable SET RITRWK=" & txtRwk & " WHERE " _
          & "RITREF='" & sCompTag & "' AND RITITM=" & str(lblItem)
   clsADOCon.ExecuteSQL sSql
   
   
End Sub

Private Sub txtScr_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdLst_Click
   If KeyCode = vbKeyPageDown Then cmdNxt_Click
   
End Sub

Private Sub txtScr_LostFocus()
   txtScr = CheckLen(txtScr, 10)
   txtScr = Format(txtScr, ES_QuantityDataFormat)
   If Val(txtScr) > cQuantity Then
      txtScr = Format(cQuantity, ES_QuantityDataFormat)
      txtRwk = Format(0, "#.000")
   End If
   If Val(txtScr) > (Val(txtRwk) + cQuantity) Then
      txtScr = Format(Val(txtQty) - Val(txtRwk), ES_QuantityDataFormat)
   End If
   If (Val(txtScr) + Val(txtRwk)) > cQuantity Then
      txtScr = Format(Val(txtQty) - Val(txtRwk), ES_QuantityDataFormat)
   End If
   On Error Resume Next
   sSql = "UPDATE RjitTable SET RITSCRP=" & txtScr & " WHERE " _
          & "RITREF='" & sCompTag & "' AND RITITM=" & str(lblItem)
   clsADOCon.ExecuteSQL sSql
   
   
End Sub

Private Sub txtTst_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdLst_Click
   If KeyCode = vbKeyPageDown Then cmdNxt_Click
   
End Sub


Private Sub txtTst_LostFocus()
   On Error Resume Next
   txtTst = CheckLen(txtTst, 1020)
   txtTst = StrCase(txtTst, ES_FIRSTWORD)
   sSql = "UPDATE RjitTable SET RITINVS='" & Trim(txtTst) & "' WHERE " _
          & "RITREF='" & sCompTag & "' AND RITITM=" & Trim(lblItem)
   clsADOCon.ExecuteSQL sSql
   
End Sub



Private Sub GetCurrentItem()
   Dim RdoBlob As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT RITREF,RITITM,RITCHARCODE,RITDATE," _
          & "RITQTY,RITRWK,RITSCRP,RITDESC," _
          & "RITDISPCODE,RITRESPCODE,RITCAUSE " _
          & "FROM RjitTable WHERE RITREF='" _
          & sCompTag & "' AND RITITM=" & str(iItemIndex)
   bSqlRows = clsADOCon.GetDataSet(sSql, AdoItems, ES_KEYSET)
   If bSqlRows Then
      With AdoItems
         lblItem = Trim(str(!RITITM))
         txtDis = "" & Trim(!RITDESC)
         If Not IsNull(!RITDATE) Then
            txtDte = Format(!RITDATE, "mm/dd/yy")
         Else
            txtDte = ""
         End If
         If !RITITM = iItemCount Then cmdNxt.Enabled = False
         If !RITITM = 1 Then cmdLst.Enabled = False
         txtQty = Format(!RITQTY, ES_QuantityDataFormat)
         txtRwk = Format(!RITRWK, ES_QuantityDataFormat)
         txtScr = Format(!RITSCRP, ES_QuantityDataFormat)
         cQuantity = Val(txtQty)
         cmbCde = "" & Trim(!RITCHARCODE)
         '10/14/02
         'cmbRes = "" & Trim(!RITRESPCODE)
         cmbCas = "" & Trim(!RITCAUSE)
         cmbDis = "" & Trim(!RITDISPCODE)
         ClearResultSet AdoItems
         GetCharaCode
         GetDispCode
         GetRespCode
      End With
      sSql = "SELECT RITREF,RITITM,RITINVS " _
             & "FROM RjitTable WHERE RITREF='" _
             & sCompTag & "' AND RITITM=" & str(iItemIndex)
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoBlob, ES_KEYSET)
      If bSqlRows Then
         With RdoBlob
            txtTst = "" & Trim(!RITINVS)
            ClearResultSet RdoBlob
         End With
      End If
      sSql = "SELECT RITREF,RITITM,RITDISP " _
             & "FROM RjitTable WHERE RITREF='" _
             & sCompTag & "' AND RITITM=" & str(iItemIndex)
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoBlob, ES_KEYSET)
      If bSqlRows Then
         With RdoBlob
            txtDip = "" & Trim(!RITDISP)
            ClearResultSet RdoBlob
         End With
      End If
      sSql = "SELECT RITREF,RITITM,RITCORA " _
             & "FROM RjitTable WHERE RITREF='" _
             & sCompTag & "' AND RITITM=" & str(iItemIndex)
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoBlob, ES_KEYSET)
      If bSqlRows Then
         With RdoBlob
            txtCor = "" & Trim(!RITCORA)
            ClearResultSet RdoBlob
         End With
      End If
   End If
   Set RdoBlob = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getcurrentit"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub DeleteItem()
   Dim bResponse As Byte
   
   bResponse = MsgBox("Really Delete Item " & lblItem & "?", ES_NOQUESTION, Caption)
   If bResponse = vbNo Then Exit Sub
   On Error GoTo DiaErr1
   MouseCursor 13
   sSql = "DELETE FROM RjitTable WHERE "
   sSql = sSql & "RITREF='" & sCompTag & "' AND RITITM=" & str(lblItem)
   clsADOCon.ExecuteSQL sSql
   MsgBox "Item Deleted.", vbInformation, Caption
   If iItemCount = 1 Then
      Unload Me
   Else
      If iItemIndex > 1 Then
         cmdLst_Click
      Else
         cmdNxt_Click
      End If
   End If
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   CurrError.Description = Err.Description
   MouseCursor 0
   MsgBox CurrError.Description & vbCr & " Can't Delete Record.", vbExclamation, Caption
   
End Sub

Private Sub GetCharaCode()
   Dim RdoCha As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT CDEREF,CDENUM,CDEDESC FROM RjcdTable " _
          & "WHERE CDEREF='" & Compress(cmbCde) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCha, ES_FORWARD)
   If bSqlRows Then
      With RdoCha
         cmbCde = "" & Trim(!CDENUM)
         lblCde = "" & Trim(!CDEDESC)
         ClearResultSet RdoCha
      End With
   Else
      If Len(Trim(cmbCde)) > 0 Then lblCde = "*** Characteristic Wasn't Found ***"
   End If
   Set RdoCha = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getcharco"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub GetRespCode()
   Dim RdoRsp As ADODB.Recordset
   If Trim(cmbRes) = "" Then
      lblRes = ""
      Exit Sub
   End If
   On Error GoTo DiaErr1
   sSql = "SELECT RESREF,RESNUM,RESDESC FROM RjrsTable " _
          & "WHERE RESREF='" & Compress(cmbRes) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRsp, ES_FORWARD)
   If bSqlRows Then
      With RdoRsp
         cmbRes = "" & Trim(!RESNUM)
         lblRes = "" & Trim(!RESDESC)
         ClearResultSet RdoRsp
      End With
   Else
      lblRes = "*** Responsibility Wasn't Found ***"
   End If
   Set RdoRsp = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getrespco"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub GetDispCode()
   Dim RdoDos As ADODB.Recordset
   If Trim(cmbDis) = "" Then
      lblDis = ""
      Exit Sub
   End If
   On Error GoTo DiaErr1
   sSql = "SELECT DISREF,DISNUM,DISDESC FROM RjdsTable " _
          & "WHERE DISREF='" & Compress(cmbDis) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoDos, ES_FORWARD)
   If bSqlRows Then
      With RdoDos
         cmbDis = "" & Trim(!DISNUM)
         lblDis = "" & Trim(!DISDESC)
         ClearResultSet RdoDos
      End With
   Else
      lblDis = "*** Disposition Wasn't Found ***"
   End If
   Set RdoDos = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getdispco"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Private Sub GetCause()
   Dim RdoRsp As ADODB.Recordset
   If Trim(cmbCas) = "" Then
      lblRes = ""
      Exit Sub
   End If
   On Error GoTo DiaErr1
   sSql = "SELECT CAUSEREF,CAUSENUM,CAUSEDESC FROM RjcaTable " _
          & "WHERE CAUSEREF='" & Compress(cmbCas) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRsp, ES_FORWARD)
   If bSqlRows Then
      With RdoRsp
         cmbCas = "" & Trim(!CAUSENUM)
         lblRes = "" & Trim(!CAUSEDESC)
         ClearResultSet RdoRsp
      End With
   Else
      lblRes = "*** Cause Code Wasn't Found ***"
   End If
   Set RdoRsp = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getcause"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub
