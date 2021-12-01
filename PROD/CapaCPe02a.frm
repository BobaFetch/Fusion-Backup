VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form CapaCPe02a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Shops"
   ClientHeight    =   4995
   ClientLeft      =   2205
   ClientTop       =   1410
   ClientWidth     =   5850
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   HelpContextID   =   4204
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4995
   ScaleWidth      =   5850
   ShowInTaskbar   =   0   'False
   Begin VB.Frame z2 
      Height          =   30
      Left            =   120
      TabIndex        =   31
      Top             =   1320
      Width           =   5652
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "CapaCPe02a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   30
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.TextBox txtSrte 
      Height          =   285
      Left            =   1560
      TabIndex        =   8
      Tag             =   "1"
      Text            =   " "
      ToolTipText     =   "Default Shop Rate For Labor Costs"
      Top             =   3240
      Width           =   825
   End
   Begin VB.ComboBox cmbAct 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1560
      Sorted          =   -1  'True
      TabIndex        =   3
      Tag             =   "3"
      Top             =   1780
      Width           =   1935
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5040
      Top             =   5160
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4995
      FormDesignWidth =   5850
   End
   Begin VB.CheckBox optSrv 
      Alignment       =   1  'Right Justify
      Caption         =   "Shop For "
      Height          =   195
      Left            =   240
      TabIndex        =   11
      ToolTipText     =   "Not Used. See Work Centers"
      Top             =   4800
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtFix 
      Height          =   285
      Left            =   1560
      TabIndex        =   10
      Tag             =   "1"
      Text            =   " "
      ToolTipText     =   "Fixed Dollar Amount"
      Top             =   4080
      Width           =   825
   End
   Begin VB.TextBox txtRte 
      Height          =   285
      Left            =   1560
      TabIndex        =   9
      Tag             =   "1"
      Text            =   " "
      Top             =   3600
      Width           =   825
   End
   Begin VB.TextBox txtUnt 
      Height          =   285
      Left            =   4080
      TabIndex        =   7
      Tag             =   "1"
      Text            =   " "
      ToolTipText     =   "Default Hours Of Unit Or Cycle Time"
      Top             =   2880
      Width           =   825
   End
   Begin VB.TextBox txtSet 
      Height          =   285
      Left            =   1560
      TabIndex        =   6
      Tag             =   "1"
      Text            =   " "
      ToolTipText     =   "Default Hours Of Setup Time"
      Top             =   2880
      Width           =   825
   End
   Begin VB.TextBox txtMdy 
      Height          =   285
      Left            =   4080
      TabIndex        =   5
      Tag             =   "1"
      Text            =   " "
      ToolTipText     =   "Default Hours Of Move Time"
      Top             =   2520
      Width           =   825
   End
   Begin VB.TextBox txtQdy 
      Height          =   285
      Left            =   1560
      TabIndex        =   4
      Tag             =   "1"
      Text            =   " "
      ToolTipText     =   "Default Hours Of Queue Time"
      Top             =   2520
      Width           =   825
   End
   Begin VB.TextBox txtEst 
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Tag             =   "1"
      Text            =   " "
      ToolTipText     =   "Default Estimating Rate"
      Top             =   1440
      Width           =   825
   End
   Begin VB.CheckBox optDef 
      Alignment       =   1  'Right Justify
      Caption         =   "Default Shop"
      Height          =   195
      Left            =   180
      TabIndex        =   12
      ToolTipText     =   "Default Shop For This Work Station"
      Top             =   4440
      Width           =   1575
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Close"
      Height          =   435
      Left            =   4920
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin VB.TextBox txtDsc 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Tag             =   "2"
      Text            =   " "
      Top             =   990
      Width           =   3075
   End
   Begin VB.ComboBox cmbShp 
      Height          =   315
      Left            =   1560
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Enter New (12 Char) Or Select From List"
      Top             =   630
      Width           =   1815
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Shop Rate"
      Height          =   285
      Index           =   14
      Left            =   180
      TabIndex        =   29
      Top             =   3240
      Width           =   1275
   End
   Begin VB.Label lblActdsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "   "
      Height          =   285
      Left            =   1560
      TabIndex        =   28
      Top             =   2160
      Width           =   3135
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Outside Services (9/8/05)"
      Height          =   285
      Index           =   13
      Left            =   2520
      TabIndex        =   27
      Top             =   4800
      Visible         =   0   'False
      Width           =   2355
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "For This Workstation"
      Height          =   285
      Index           =   12
      Left            =   2520
      TabIndex        =   26
      Top             =   4440
      Width           =   1815
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "If Fixed Hourly Rate"
      Height          =   285
      Index           =   11
      Left            =   2520
      TabIndex        =   25
      Top             =   4080
      Width           =   2805
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Overhead Rate"
      Height          =   285
      Index           =   10
      Left            =   180
      TabIndex        =   24
      Top             =   4080
      Width           =   1395
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "As Applied As A Percentage Of Employee Rate"
      Height          =   375
      Index           =   9
      Left            =   2520
      TabIndex        =   23
      Top             =   3600
      UseMnemonic     =   0   'False
      Width           =   3285
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Overhead "
      Height          =   285
      Index           =   8
      Left            =   180
      TabIndex        =   22
      Top             =   3600
      Width           =   1275
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Unit/CycleTime"
      Height          =   285
      Index           =   7
      Left            =   2700
      TabIndex        =   21
      Top             =   2880
      Width           =   1275
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Setup Time"
      Height          =   285
      Index           =   6
      Left            =   180
      TabIndex        =   20
      Top             =   2880
      Width           =   1275
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Move Hours"
      Height          =   285
      Index           =   5
      Left            =   2700
      TabIndex        =   19
      Top             =   2520
      Width           =   1275
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Queue Hours"
      Height          =   285
      Index           =   4
      Left            =   180
      TabIndex        =   18
      Top             =   2520
      Width           =   1305
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Shop Labor Acct"
      Height          =   285
      Index           =   3
      Left            =   180
      TabIndex        =   17
      Top             =   1800
      Width           =   1395
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Estimating Rate"
      Height          =   285
      Index           =   2
      Left            =   180
      TabIndex        =   16
      Top             =   1440
      Width           =   1395
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   285
      Index           =   1
      Left            =   180
      TabIndex        =   15
      Top             =   990
      Width           =   1125
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Add/ Revise Shop"
      Height          =   375
      Index           =   0
      Left            =   180
      TabIndex        =   14
      Top             =   540
      Width           =   1185
   End
End
Attribute VB_Name = "CapaCPe02a"
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
Dim AdoQry As ADODB.Command
Dim AdoParameter As ADODB.Parameter
Dim AdoShp As ADODB.Recordset

Dim bCancel As Byte
Dim bGoodShop As Byte
Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub ManageBoxes(bEnabled As Boolean)
   Dim iList As Integer
   For iList = 0 To Controls.Count - 1
      If Err > 0 And (TypeOf Controls(iList) Is TextBox Or _
                      TypeOf Controls(iList) Is ComboBox Or TypeOf Controls(iList) Is MaskEdBox) Then
         If Controls(iList).TabIndex > 2 Then Controls(iList).Enabled = bEnabled
      End If
   Next
   optSrv.Enabled = bEnabled
   optDef.Enabled = bEnabled
   cmbAct.Enabled = False
   
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   ES_TimeFormat = GetTimeFormat()
   
End Sub

Private Sub cmbAct_Click()
   FindAccount Me
   
End Sub

Private Sub cmbAct_LostFocus()
   Dim b As Byte
   Dim iList As Integer
   Dim sAccount As String
   cmbAct = CheckLen(cmbAct, 12)
   If Len(cmbAct) Then
      For iList = 0 To cmbAct.ListCount - 1
         If cmbAct = cmbAct.List(iList) Then b = 1
      Next
      On Error Resume Next
      If b = 0 Then
         Beep
         cmbAct = "" & Trim(AdoShp!SHPACCT)
      End If
      FindAccount Me
      sAccount = Compress(cmbAct)
   End If
   If bGoodShop Then
      AdoShp!SHPACCT = "" & sAccount
      AdoShp.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub cmbShp_Click()
   bGoodShop = GetShop(False)
   
End Sub

Private Sub cmbShp_LostFocus()
   cmbShp = CheckLen(cmbShp, 12)
   If bCancel = 1 Then Exit Sub
   If Len(cmbShp) = 0 Then
      On Error Resume Next
      cmdCan.SetFocus
      Exit Sub
   End If
   If Len(cmbShp) > 0 Then
      bGoodShop = GetShop(True)
   Else
      optDef.Value = vbUnchecked
      bGoodShop = False
      On Error Resume Next
      cmdCan.SetFocus
   End If
   If Not bGoodShop Then AddShop
   
End Sub


Private Sub cmdCan_Click()
   txtDsc_LostFocus
   Unload Me
   
End Sub


Private Sub cmdCan_LostFocus()
   Set CapaCPe02a = Nothing
   
End Sub



Private Sub FillCombo()
   On Error GoTo DiaErr1
   sSql = "Qry_FillShops"
   LoadComboBox cmbShp
   If cmbShp.ListCount > 0 Then cmbShp = cmbShp.List(0)
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Function GetShop(bOpen As Byte) As Byte
   Dim sShop As String
   sShop = Compress(cmbShp)
   
   GetShop = False
   On Error GoTo DiaErr1
   'RdoQry.RowsetSize = 1
   'RdoQry(0) = sShop
   AdoQry.Parameters(0).Value = sShop
   bSqlRows = clsADOCon.GetQuerySet(AdoShp, AdoQry, ES_KEYSET, True)
   If bSqlRows Then
      With AdoShp
         GetShop = True
         cmbShp = "" & Trim(!SHPNUM)
         txtDsc = "" & Trim(!SHPDESC)
         txtEst = Format(0 + !SHPESTRATE, ES_QuantityDataFormat)
         cmbAct = "" & Trim(!SHPACCT)
         FindAccount Me
         If cmbAct = "" Then lblActdsc = ""
         If cmbAct.ListCount > 0 Then cmbAct.Enabled = True
         txtQdy = Format(!SHPQHRS, "##0.000")
         txtMdy = Format(!SHPMHRS, "##0.000")
         txtSet = Format(!SHPSUHRS, "##0.000")
         txtUnt = Format(!SHPUNITHRS, ES_TimeFormat)
         txtSrte = Format(!SHPRATE, "##0.000")
         txtFix = Format(!SHPOHTOTAL, ES_QuantityDataFormat)
         txtRte = Format(!SHPOHRATE, ES_QuantityDataFormat)
         optSrv.Value = !SHPSERVICE
         If cUR.CurrentShop = Trim(cmbShp) Then
            optDef.Value = vbChecked
         Else
            optDef.Value = vbUnchecked
         End If
         If bOpen Then
            txtEst.Enabled = True
            If cmbAct.ListCount > 0 Then cmbAct.Enabled = True
         End If
         ManageBoxes True
      End With
   Else
      GetShop = False
      ManageBoxes False
      On Error Resume Next
      cmbShp.SetFocus
   End If
   Exit Function
   
DiaErr1:
   sProcName = "getshop"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   bCancel = 1
   
End Sub


Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 4204
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub



Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      ManageBoxes False
      FillAccounts
      FillCombo
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   On Error Resume Next
   sSql = "SELECT TOP 1 * FROM ShopTable WHERE SHPREF= ? "
   'Set RdoQry = RdoCon.CreateQuery("", sSql)
   'RdoQry.MaxRows = 1
   
   Set AdoQry = New ADODB.Command
   AdoQry.CommandText = sSql
   
   Set AdoParameter = New ADODB.Parameter
   AdoParameter.Type = adChar
   AdoParameter.SIZE = 12
   
   AdoQry.Parameters.Append AdoParameter
   
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set AdoParameter = Nothing
   Set AdoQry = Nothing
   Set AdoShp = Nothing
   Set CapaCPe02a = Nothing
   
End Sub



Private Sub AddShop()
   Dim sNewShop As String
   Dim bResponse As Byte
   
   bResponse = MsgBox(cmbShp & " Wasn't Found. Add It?", ES_YESQUESTION, Caption)
   If bResponse = vbNo Then
      bGoodShop = False
      On Error Resume Next
      cmbShp = cmbShp.List(0)
      cmbShp.SetFocus
      Width = Width + 10
      Exit Sub
   End If
   bResponse = IllegalCharacters(cmbShp)
   If bResponse > 0 Then
      MsgBox "The Shop Contains An Illegal " & Chr$(bResponse) & ".", _
         vbExclamation, Caption
      Exit Sub
   End If
   MouseCursor 11
   sNewShop = Compress(cmbShp)
   
   On Error Resume Next
   'AdoShp.Close
   On Error GoTo DiaErr1
   sSql = "INSERT INTO ShopTable (SHPREF, SHPNUM) VALUES ('" & sNewShop & "','" & cmbShp & "') "
   clsADOCon.ExecuteSQL sSql
   
   'sSql = "Select * FROM ShopTable"
   
   'Set AdoShp = RdoCon.OpenResultset(sSql, rdOpenKeyset, rdConcurRowVer)
   'clsAdoCon.begintrans
   'AdoShp.AddNew
   'AdoShp!SHPREF = sNewShop
   'AdoShp!SHPNUM = cmbShp
   'AdoShp.Update
   'clsAdoCon.CommitTrans
   On Error Resume Next
   AddComboStr cmbShp.hwnd, cmbShp
   'AdoShp.Close
   MouseCursor 0
   bGoodShop = GetShop(True)
   txtDsc.SetFocus
   SysMsg cmbShp & " Added.", True
   Exit Sub
   
DiaErr1:
   sProcName = "addshop"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   Resume DiaErr2
DiaErr2:
   MouseCursor 0
   On Error Resume Next
   'clsAdoCon.RollbackTrans
   MsgBox CurrError.Description & vbCr & "Couldn't Add Shop.", vbExclamation, Caption
   DoModuleErrors Me
   
End Sub

Private Sub lblActdsc_Change()
   If Trim(lblActdsc) = "*** Account Wasn't Found ***" Then
      lblActdsc.ForeColor = ES_RED
   Else
      lblActdsc.ForeColor = Es_TextForeColor
   End If
   
End Sub

Private Sub optDef_Click()
   SaveSetting "Esi2000", "Current", "Shop", cUR.CurrentShop
   
End Sub

Private Sub optSrv_Click()
   If bGoodShop Then
      On Error Resume Next
      AdoShp!SHPSERVICE = optSrv.Value
      AdoShp.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub



Private Sub txtDsc_LostFocus()
   txtDsc = CheckLen(txtDsc, 30)
   txtDsc = StrCase(txtDsc)
   If bGoodShop Then
      On Error Resume Next
      AdoShp!SHPDESC = "" & txtDsc
      AdoShp.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtEst_LostFocus()
   txtEst = CheckLen(txtEst, 7)
   txtEst = Format(Abs(Val(txtEst)), ES_QuantityDataFormat)
   If bGoodShop Then
      On Error Resume Next
      AdoShp!SHPESTRATE = Val(txtEst)
      AdoShp.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub

Private Sub txtFix_LostFocus()
   txtFix = CheckLen(txtFix, 7)
   If Val(txtFix) > 100 Then txtFix = "100.000"
   If bGoodShop Then
      On Error Resume Next
      AdoShp!SHPOHTOTAL = Val(txtFix)
      AdoShp.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub

Private Sub txtMdy_LostFocus()
   txtMdy = CheckLen(txtMdy, 7)
   txtMdy = Format(Abs(Val(txtMdy)), ES_QuantityDataFormat)
   If bGoodShop Then
      On Error Resume Next
      AdoShp!SHPMHRS = Val(txtMdy)
      AdoShp.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub

Private Sub txtQdy_LostFocus()
   txtQdy = CheckLen(txtQdy, 7)
   txtQdy = Format(Abs(Val(txtQdy)), ES_QuantityDataFormat)
   If bGoodShop Then
      On Error Resume Next
      AdoShp!SHPQHRS = Val(txtQdy)
      AdoShp.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub

Private Sub txtRte_LostFocus()
   txtRte = CheckLen(txtRte, 7)
   txtRte = Format(Abs(Val(txtRte)), ES_QuantityDataFormat)
   If bGoodShop Then
      On Error Resume Next
      AdoShp!SHPOHRATE = Val(txtRte)
      AdoShp.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub

Private Sub txtSet_LostFocus()
   txtSet = CheckLen(txtSet, 7)
   txtSet = Format(Abs(Val(txtSet)), ES_QuantityDataFormat)
   If bGoodShop Then
      On Error Resume Next
      AdoShp!SHPSUHRS = Val(txtSet)
      AdoShp.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub

Private Sub txtSrte_LostFocus()
   txtSrte = CheckLen(txtSrte, 7)
   txtSrte = Format(Abs(Val(txtSrte)), ES_QuantityDataFormat)
   If bGoodShop Then
      On Error Resume Next
      AdoShp!SHPRATE = Val(txtSrte)
      AdoShp.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtUnt_LostFocus()
   txtUnt = CheckLen(txtUnt, 8)
   txtUnt = Format(Abs(Val(txtUnt)), ES_TimeFormat)
   If bGoodShop Then
      On Error Resume Next
      AdoShp!SHPUNITHRS = Val(txtUnt)
      AdoShp.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub



Private Sub FillAccounts()
   On Error GoTo DiaErr1
   sSql = "Qry_FillLowAccounts"
   LoadComboBox cmbAct
   If cmbAct.ListCount > 0 Then
      cmbAct = cmbAct.List(0)
      cmbAct.Enabled = True
      FindAccount Me
   Else
      cmbAct = "No Accounts."
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "fillaccou"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub
