VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form EstiESe02b 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Estimate Routing "
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6375
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "EstiESe02b.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   30
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdComments 
      DisabledPicture =   "EstiESe02b.frx":07AE
      DownPicture     =   "EstiESe02b.frx":1120
      Height          =   350
      Index           =   1
      Left            =   5280
      Picture         =   "EstiESe02b.frx":1A92
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "Standard Comments"
      Top             =   4140
      Width           =   350
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "&Delete"
      Enabled         =   0   'False
      Height          =   285
      Left            =   5400
      TabIndex        =   28
      TabStop         =   0   'False
      ToolTipText     =   "Delete The Current Entry"
      Top             =   600
      Width           =   885
   End
   Begin VB.TextBox txtRte 
      Height          =   285
      Left            =   4800
      TabIndex        =   6
      Tag             =   "1"
      ToolTipText     =   "Labor Rate"
      Top             =   3300
      Width           =   825
   End
   Begin VB.CommandButton cmdUpd 
      Caption         =   "&Apply"
      Height          =   285
      Left            =   5280
      TabIndex        =   20
      TabStop         =   0   'False
      ToolTipText     =   "Update Current Operation"
      Top             =   5340
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&Add"
      Height          =   285
      Left            =   5400
      TabIndex        =   19
      TabStop         =   0   'False
      ToolTipText     =   "Add A New Operation"
      Top             =   2820
      Width           =   885
   End
   Begin VB.TextBox txtCmt 
      Height          =   765
      Left            =   1440
      MultiLine       =   -1  'True
      TabIndex        =   9
      Tag             =   "9"
      ToolTipText     =   "Notes: 255 Char Max"
      Top             =   4140
      Width           =   3705
   End
   Begin VB.TextBox txtQdy 
      Height          =   285
      Left            =   1440
      TabIndex        =   7
      Tag             =   "1"
      ToolTipText     =   "Queue Hours"
      Top             =   3660
      Width           =   825
   End
   Begin VB.TextBox txtMdy 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   3240
      TabIndex        =   8
      Tag             =   "1"
      ToolTipText     =   "Move Hours"
      Top             =   3660
      Width           =   825
   End
   Begin VB.TextBox txtSet 
      Height          =   285
      Left            =   1440
      TabIndex        =   4
      Tag             =   "1"
      ToolTipText     =   "Lot Setup Hours"
      Top             =   3300
      Width           =   825
   End
   Begin VB.TextBox txtUnt 
      Height          =   285
      Left            =   3240
      TabIndex        =   5
      Tag             =   "1"
      ToolTipText     =   "Unit Hours"
      Top             =   3300
      Width           =   825
   End
   Begin MSFlexGridLib.MSFlexGrid Grd 
      Height          =   2295
      Left            =   300
      TabIndex        =   0
      ToolTipText     =   "Click To Select Or Scroll And Press Enter"
      Top             =   120
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   4048
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      FocusRect       =   2
      ScrollBars      =   2
      SelectionMode   =   1
   End
   Begin VB.TextBox txtOpn 
      ForeColor       =   &H00400000&
      Height          =   315
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Tag             =   "1"
      ToolTipText     =   "Operation Number (Locked)"
      Top             =   2820
      Width           =   435
   End
   Begin VB.ComboBox cmbShp 
      Height          =   315
      Left            =   1080
      Sorted          =   -1  'True
      TabIndex        =   2
      Tag             =   "3"
      Top             =   2820
      Width           =   1815
   End
   Begin VB.ComboBox cmbWcn 
      Height          =   315
      Left            =   3000
      TabIndex        =   3
      Tag             =   "3"
      Top             =   2820
      Width           =   1815
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5400
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5880
      Top             =   4500
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   5985
      FormDesignWidth =   6375
   End
   Begin VB.Label z1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      Height          =   285
      Index           =   10
      Left            =   5640
      TabIndex        =   27
      Top             =   3660
      Width           =   405
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "FOH"
      Height          =   285
      Index           =   7
      Left            =   4200
      TabIndex        =   26
      Top             =   3660
      Width           =   645
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Rate"
      Height          =   285
      Index           =   4
      Left            =   4200
      TabIndex        =   25
      Top             =   3300
      Width           =   645
   End
   Begin VB.Label lblEdit 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "No Current Operation"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   600
      TabIndex        =   24
      Top             =   5100
      Width           =   4455
   End
   Begin VB.Label lblFoh 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   4800
      TabIndex        =   23
      ToolTipText     =   "Factory Overhead Rate"
      Top             =   3660
      Width           =   825
   End
   Begin VB.Label lblRate 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5280
      TabIndex        =   22
      ToolTipText     =   "Labor Rate"
      Top             =   4980
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblBid 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   120
      TabIndex        =   21
      ToolTipText     =   "Total Services"
      Top             =   5580
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Op No "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   480
      TabIndex        =   18
      Top             =   2580
      Width           =   675
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Shop                               "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   1080
      TabIndex        =   17
      Top             =   2580
      Width           =   2055
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Work Center                     "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   3000
      TabIndex        =   16
      Top             =   2580
      Width           =   1935
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Notes:"
      Height          =   285
      Index           =   0
      Left            =   480
      TabIndex        =   15
      Top             =   4140
      Width           =   795
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Queue Hrs"
      Height          =   285
      Index           =   5
      Left            =   480
      TabIndex        =   14
      Top             =   3660
      Width           =   1035
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Move Hrs"
      Height          =   285
      Index           =   6
      Left            =   2400
      TabIndex        =   13
      Top             =   3660
      Width           =   885
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Setup"
      Height          =   285
      Index           =   8
      Left            =   480
      TabIndex        =   12
      Top             =   3300
      Width           =   1065
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Unit"
      Height          =   285
      Index           =   9
      Left            =   2400
      TabIndex        =   11
      Top             =   3300
      Width           =   645
   End
End
Attribute VB_Name = "EstiESe02b"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'11/2/04 Reworked Grid and Work Centers
'3/31/06 Added Comments Selection
Option Explicit
Dim RdoOpn As ADODB.Recordset
Dim bCanceled As Byte
Dim bOnLoad As Byte
Dim bGoodBid As Byte

Dim bUseCenter As Byte
Dim iGridIndex As Integer
Dim lBidNo As Long

Dim cDefaultFoh As Currency
Dim cDefaultRate As Currency

Dim cCurrentFoh As Currency
Dim cCurrentRate As Currency

Dim sEstimator As String
Dim sOldShop As String
Dim sOldCenter As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Const LABORCOL_OpNo = 0
Private Const LABORCOL_Shop = 1
Private Const LABORCOL_WC = 2
Private Const LABORCOL_Hours = 3

Private GettingExistingOperation As Boolean

Private Sub GetRates()
   Dim RdoPar As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT EstFactoryOverHead,EstLaborRate," _
          & "EstUseWCOverhead From Preferences WHERE " _
          & "PreRecord=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPar, ES_FORWARD)
   If bSqlRows Then
      With RdoPar
         cDefaultFoh = Format(!EstFactoryOverHead / 100, ES_QuantityDataFormat)
         cDefaultRate = Format(!EstLaborRate, ES_QuantityDataFormat)
         bUseCenter = !EstUseWCOverhead
         ClearResultSet RdoPar
         txtRte = Format(cDefaultRate, ES_QuantityDataFormat)
         lblFoh = Format(cDefaultFoh * 100, ES_QuantityDataFormat)
      End With
   Else
      MouseCursor 0
      txtRte = "0.000"
      lblFoh = "0.000"
      MsgBox "There Are No Default Rates.", _
         vbInformation, Caption
   End If
   sEstimator = GetSetting("Esi2000", "EsiEngr", "Estimator", sEstimator)
   sCurrEstimator = GetSetting("Esi2000", "EsiEngr", "Estimator", sCurrEstimator)
   Set RdoPar = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getrates"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub FillWorkCenters(SelectKey As String)
   'if SelectKey <> "" then select the key with this value
   cmbWcn.Clear
   On Error GoTo DiaErr1
   sSql = "Qry_FillWorkCenters '" & Compress(cmbShp) & "'"
   LoadComboBoxAndSelect cmbWcn, SelectKey
   If cmbWcn.ListCount > 0 Then
'      If sOldShop = cmbShp Then
'         cmbWcn = sOldCenter
'      Else
'         cmbWcn = cmbWcn.List(0)
'      End If
      
      If SelectKey <> "" Then
         cmbWcn = SelectKey
      Else
         cmbWcn = cmbWcn.List(0)
      End If
      
      GetCenterInfo
   End If
   If Grd.Row > 0 Then
      Grd.Col = 2
      Grd.Text = cmbWcn
   End If
   sOldShop = cmbShp
   sOldCenter = cmbWcn
   Exit Sub
   
DiaErr1:
   sProcName = "fillworkc"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub cmbShp_Click()
   GetShopInfo
   'Grd.Col = 1
   'If Grd.Row > 0 Then Grd.Text = cmbShp
   'If sOldShop <> cmbShp Then FillWorkCenters
End Sub


Private Sub cmbShp_LostFocus()
   Dim b As Byte
   Dim iList As Integer
   If bCanceled = 1 Then
      bCanceled = 0
      Exit Sub
   End If
   cmbShp = CheckLen(cmbShp, 12)
   For iList = 0 To cmbShp.ListCount - 1
      If Trim(cmbShp) = Trim(cmbShp.List(iList)) Then b = 1
   Next
   If b = 0 Then
      Beep
      cmbShp = cUR.CurrentShop
   End If
   GetShopInfo
   FillWorkCenters ""
   sOldShop = cmbShp
   Grd.Col = 1
   Grd.Text = cmbShp
   If bGoodBid = 1 Then
      'On Error Resume Next
      'RdoOpn.Edit
      RdoOpn!BIDRTESHOP = Compress(cmbShp)
      RdoOpn.Update
   End If
   
End Sub


Private Sub cmbWcn_Click()
   GetCenterInfo
   'Grd.Col = 2
   'If Grd.Row > 0 Then Grd.Text = cmbWcn
   'If Grd.Row > 0 Then Grd.TextMatrix(Grd.Row, 2) = cmbWcn
   
End Sub

Private Sub cmbWcn_LostFocus()
   'if shop/wc unchanged, no action required
   If cmbShp = Grd.TextMatrix(Grd.Row, LABORCOL_Shop) _
      And cmbWcn = Grd.TextMatrix(Grd.Row, LABORCOL_WC) Then
      Exit Sub
   End If
   
   Dim b As Byte
   Dim iList As Integer
   If bCanceled = 1 Then
      bCanceled = 0
      Exit Sub
   End If
   cmbWcn = CheckLen(cmbWcn, 12)
   For iList = 0 To cmbWcn.ListCount - 1
      If cmbWcn = cmbWcn.List(iList) Then b = 1
   Next
   If b = 0 Then
      Beep
      If cmbWcn.ListCount > 0 Then _
         cmbWcn = cmbWcn.List(0)
   End If
   GetCenterInfo
   'sOldCenter = cmbWcn -- NO -- compare against the grid cell to see if there is a change
   Grd.Row = iGridIndex
   Grd.Col = 2
   Grd.Text = cmbWcn
   If bGoodBid = 1 Then
      'On Error Resume Next
      'RdoOpn.Edit
      RdoOpn!BIDRTECENTER = Compress(cmbWcn)
      RdoOpn.Update
   End If
   
   
End Sub


Private Sub cmbWcn_Validate(Cancel As Boolean)
   Debug.Print "validate"
End Sub

Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   bCanceled = 1
   
End Sub


Private Sub cmdComments_Click(Index As Integer)
   txtCmt.SetFocus
   'See List For Index
   SysComments.lblListIndex = 7
   SysComments.Show
   cmdComments(1) = False
   
End Sub

Private Sub cmdDel_Click()
   Dim bResponse As Byte
   Dim sMsg As String
   
   sMsg = "Are You Want To Delete The Entry For " & vbCrLf _
          & "Operation " & Trim(txtOpn) & "?"
   bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
   If bResponse = vbYes Then
      'On Error Resume Next
      RdoOpn.Close
      clsADOCon.ADOErrNum = 0
      sSql = "DELETE FROM EsrtTable WHERE BIDRTEREF=" _
             & Val(lblBid) & " AND BIDRTEOPNO=" _
             & Val(txtOpn) & " "
      clsADOCon.ExecuteSQL sSql ' rdExecDirect
      If clsADOCon.ADOErrNum = 0 Then
         SysMsg "Operation Was Deleted.", True
         If Grd.Rows < 3 Then
            MsgBox "There Are No Operations Remaining.", _
               vbInformation, Caption
            Unload Me
         Else
            RewindOps
            GetOperations
         End If
      Else
         MsgBox "Could Not Delete The Operation.", _
            vbExclamation, Caption
      End If
   Else
      CancelTrans
   End If
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 3510
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub


Private Sub cmdNew_Click()
   AddANewOperation
   
End Sub

Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      MouseCursor 13
      iGridIndex = 0
      Caption = Caption & " - Estimate " & lblBid
      GetRates
      lBidNo = Val(lblBid)
      FillCombo
      If cmbWcn.ListCount = 0 Or cmbShp.ListCount = 0 Then
         MsgBox "There Are Either No Shops Or No Work Centers.", _
            vbExclamation, Caption
         Unload Me
      Else
         GetOperations
      End If
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Deactivate()
   Unload Me
   
End Sub

Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST
   Move 1000, 1000
   FormatControls
   With Grd
      .ColAlignment(0) = 0
      .ColAlignment(1) = 0
      .ColAlignment(2) = 0
      .Row = 0
      .Col = 0
      .Text = "Op No"
      .ColWidth(0) = 750
      .Col = 1
      .Text = "Shop"
      .ColWidth(1) = 1500
      .Col = 2
      .Text = "Work Center"
      .ColWidth(2) = 1500
      .Col = 3
      .Text = "Hours"
      .ColWidth(3) = 850
      .Col = 0
   End With
   bOnLoad = 1
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Dim b As Byte
   b = GetBidLabor(Compress(EstiESe02a.txtPrt), Val(EstiESe02a.cmbBid), CCur("0" & EstiESe02a.txtQty))
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   'On Error Resume Next
   Set RdoOpn = Nothing
   Set EstiESe02b = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   ES_TimeFormat = GetTimeFormat()
   txtOpn = "000"
   txtSet = "0.000"
   txtUnt = ES_TimeFormat
   txtMdy = "0.000"
   txtQdy = "0.000"
   txtRte = "0.000"
   
End Sub

Private Sub FillCombo()
   Dim RdoCmb As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "Qry_FillShops"
   LoadComboBox cmbShp
   If cmbShp.ListCount > 0 Then
      cmbShp = cmbShp.List(0)
      GetShopInfo
      FillWorkCenters ""
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Private Sub grd_Click()
   'Grd.Col = 0
   iGridIndex = Grd.Row
   txtOpn = Grd.TextMatrix(Grd.Row, LABORCOL_OpNo)
   If Val(txtOpn) > 0 Then bGoodBid = GetThisOperation()
   
End Sub

Private Sub Grd_EnterCell()
   If iGridIndex <> Grd.Row Then
      iGridIndex = Grd.Row
      grd_Click
   End If
End Sub

Private Sub Grd_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Grd.Col = 0
      iGridIndex = Grd.Row
      txtOpn = Grd.Text
      If Val(txtOpn) > 0 Then bGoodBid = GetThisOperation()
   End If
   
End Sub


Private Sub Grd_KeyUp(KeyCode As Integer, Shift As Integer)
   iGridIndex = Grd.Row
   
End Sub


Private Sub Grd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   'bCanceled = 1
   
End Sub


Private Sub txtCmt_LostFocus()
'   txtCmt = CheckLen(txtCmt, 255)
'   If bCanceled = 1 Then Exit Sub
'   txtCmt = StrCase(txtCmt, ES_FIRSTWORD)
'   If bGoodBid = 1 Then
'      'On Error Resume Next
'      'RdoOpn.Edit
'      RdoOpn!BIDRTENOTES = txtCmt
'      RdoOpn.Update
'   End If
   
End Sub

Private Sub txtCmt_Validate(Cancel As Boolean)
   txtCmt = CheckLen(txtCmt, 255)
   If bCanceled = 1 Then Exit Sub
   txtCmt = StrCase(txtCmt, ES_FIRSTWORD)
   If bGoodBid = 1 Then
      'On Error Resume Next
      'RdoOpn.Edit
      RdoOpn!BIDRTENOTES = txtCmt
      RdoOpn.Update
   End If
End Sub

Private Sub txtMdy_LostFocus()
'   txtMdy = CheckLen(txtMdy, 7)
'   txtMdy = Format(Abs(Val(txtMdy)), ES_QuantityDataFormat)
'   If bCanceled = 1 Then Exit Sub
'   If bGoodBid = 1 Then
'      'On Error Resume Next
'      'RdoOpn.Edit
'      RdoOpn!BIDRTEMHRS = Val(txtMdy)
'      RdoOpn.Update
'   End If
   
End Sub

Private Sub txtMdy_Validate(Cancel As Boolean)
   txtMdy = CheckLen(txtMdy, 7)
   txtMdy = Format(Abs(Val(txtMdy)), ES_QuantityDataFormat)
   If bCanceled = 1 Then Exit Sub
   If bGoodBid = 1 Then
      'RdoOpn.Edit
      RdoOpn!BIDRTEMHRS = Val(txtMdy)
      RdoOpn.Update
   End If
End Sub

Private Sub txtQdy_LostFocus()
'   txtQdy = CheckLen(txtQdy, 7)
'   txtQdy = Format(Abs(Val(txtQdy)), ES_QuantityDataFormat)
'   If bCanceled = 1 Then Exit Sub
'   If bGoodBid = 1 Then
'      'On Error Resume Next
'      'RdoOpn.Edit
'      RdoOpn!BIDRTEQHRS = Val(txtQdy)
'      RdoOpn.Update
'   End If
   
End Sub


Private Sub txtQdy_Validate(Cancel As Boolean)
   txtQdy = CheckLen(txtQdy, 7)
   txtQdy = Format(Abs(Val(txtQdy)), ES_QuantityDataFormat)
   If bCanceled = 1 Then Exit Sub
   If bGoodBid = 1 Then
      'RdoOpn.Edit
      RdoOpn!BIDRTEQHRS = Val(txtQdy)
      RdoOpn.Update
   End If
End Sub

Private Sub txtRte_LostFocus()
'   txtRte = CheckLen(txtRte, 7)
'   txtRte = Format(Abs(Val(txtRte)), ES_QuantityDataFormat)
'   If bCanceled = 1 Then Exit Sub
'   If bGoodBid = 1 Then
'      'On Error Resume Next
'      'RdoOpn.Edit
'      RdoOpn!BIDRTERATE = Val(txtRte)
'      RdoOpn!BIDRTEFOHRATE = Val(lblFoh) / 100
'      RdoOpn.Update
'   End If
'   Grd.Col = 3
'   Grd.Text = Format(Val(txtSet) + Val(txtUnt), ES_QuantityDataFormat)
'   Grd.Col = 0
   
End Sub


Private Sub txtRte_Validate(Cancel As Boolean)
   txtRte = CheckLen(txtRte, 7)
   txtRte = Format(Abs(Val(txtRte)), ES_QuantityDataFormat)
   If bCanceled = 1 Then Exit Sub
   If bGoodBid = 1 Then
      'On Error Resume Next
      'RdoOpn.Edit
      RdoOpn!BIDRTERATE = Val(txtRte)
      RdoOpn!BIDRTEFOHRATE = Val(lblFoh) / 100
      RdoOpn.Update
   End If
'   Grd.Col = 3
'   Grd.Text = Format(Val(txtSet) + Val(txtUnt), ES_QuantityDataFormat)
'   Grd.Col = 0
   Grd.TextMatrix(Grd.Row, LABORCOL_Hours) = Format(Val(txtSet) + Val(txtUnt), ES_QuantityDataFormat)
   
End Sub

Private Sub txtSet_LostFocus()
'   txtSet = CheckLen(txtSet, 7)
'   txtSet = Format(Abs(Val(txtSet)), ES_QuantityDataFormat)
'   If bCanceled = 1 Then Exit Sub
'   If bGoodBid = 1 Then
'      'On Error Resume Next
'      'RdoOpn.Edit
'      RdoOpn!BIDRTESETUP = Val(txtSet)
'      RdoOpn!BIDRTERATE = Val(txtRte)
'      RdoOpn!BIDRTEFOHRATE = Val(lblFoh) / 100
'      RdoOpn.Update
'   End If
'   Grd.Col = 3
'   Grd.Text = Format(Val(txtSet) + Val(txtUnt), ES_QuantityDataFormat)
'   Grd.Col = 0
'
End Sub


Private Sub txtSet_Validate(Cancel As Boolean)
   txtSet = CheckLen(txtSet, 7)
   txtSet = Format(Abs(Val(txtSet)), ES_QuantityDataFormat)
   If bCanceled = 1 Then Exit Sub
   If bGoodBid = 1 Then
      'On Error Resume Next
      'RdoOpn.Edit
      RdoOpn!BIDRTESETUP = Val(txtSet)
      RdoOpn!BIDRTERATE = Val(txtRte)
      RdoOpn!BIDRTEFOHRATE = Val(lblFoh) / 100
      RdoOpn.Update
   End If
'   Grd.Col = 3
'   Grd.Text = Format(Val(txtSet) + Val(txtUnt), ES_QuantityDataFormat)
'   Grd.Col = 0
   Grd.TextMatrix(Grd.Row, LABORCOL_Hours) = Format(Val(txtSet) + Val(txtUnt), ES_QuantityDataFormat)
End Sub

Private Sub txtUnt_LostFocus()
'   txtUnt = CheckLen(txtUnt, 8)
'   txtUnt = Format(Abs(Val(txtUnt)), ES_TimeFormat)
'   If bCanceled = 1 Then Exit Sub
'   If bGoodBid = 1 Then
'      'On Error Resume Next
'      'RdoOpn.Edit
'      RdoOpn!BIDRTEUNIT = Val(txtUnt)
'      RdoOpn!BIDRTERATE = Val(txtRte)
'      RdoOpn!BIDRTEFOHRATE = Val(lblFoh) / 100
'      RdoOpn.Update
'   End If
'   Grd.Col = 3
'   Grd.Text = Format(Val(txtSet) + Val(txtUnt), ES_QuantityDataFormat)
'   Grd.Col = 0
'
End Sub



Private Sub GetCenterInfo()
   Dim RdoCnt As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT WCNREF,WCNNUM,WCNDESC,WCNESTRATE," _
          & "WCNOHPCT FROM WcntTable WHERE WCNREF='" _
          & Compress(cmbWcn) & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCnt, ES_FORWARD)
   If bSqlRows Then
      With RdoCnt
         cmbWcn = "" & Trim(!WCNNUM)
         cmbWcn.ToolTipText = "" & Trim(!WCNDESC)
         If bUseCenter = 1 Then
            cCurrentRate = !WCNESTRATE
            cCurrentFoh = !WCNOHPCT / 100
         Else
            cCurrentRate = cDefaultRate
            cCurrentFoh = cDefaultFoh
         End If
         ClearResultSet RdoCnt
      End With
   Else
      cmbWcn.ToolTipText = "Select Valid Work Center From The List"
   End If
   If cCurrentRate = 0 Then
      cCurrentRate = cDefaultRate
   End If
   'If sOldCenter <> cmbWcn Or Val(txtRte) = 0 Then
   If bOnLoad = 0 And Grd.Row <> 0 Then
      If cmbShp <> Grd.TextMatrix(Grd.Row, LABORCOL_Shop) _
         Or cmbWcn <> Grd.TextMatrix(Grd.Row, LABORCOL_WC) Then
         txtRte = Format(cCurrentRate, ES_QuantityDataFormat)
      End If
   End If
   'End If
   lblFoh = Format(cCurrentFoh * 100, ES_QuantityDataFormat)
   Set RdoCnt = Nothing
   Exit Sub
      
DiaErr1:
   sProcName = "getcenterinfo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub GetShopInfo()
   Dim RdoShp As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT SHPREF,SHPNUM,SHPDESC FROM ShopTable " _
          & "WHERE SHPREF='" & Compress(cmbShp) & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoShp, ES_FORWARD)
   If bSqlRows Then
      With RdoShp
         cmbShp = "" & Trim(!SHPNUM)
         cmbShp.ToolTipText = "" & Trim(!SHPDESC)
         ClearResultSet RdoShp
      End With
   Else
      cmbShp.ToolTipText = "Select Valid Shop From The List"
   End If
   Set RdoShp = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getshopinfo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub GetOperations()
   Dim RdoOps As ADODB.Recordset
   Dim iRow As Integer
   
   On Error GoTo DiaErr1
   Grd.Rows = 2
   sSql = "SELECT BIDRTEREF,BIDRTEOPNO,BIDRTESHOP," _
          & "BIDRTECENTER,BIDRTESETUP,BIDRTEUNIT FROM " _
          & "EsrtTable WHERE BIDRTEREF=" & Val(lblBid) & " "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoOps, ES_FORWARD)
   If bSqlRows Then
      MouseCursor 13
      With RdoOps
         iRow = 0
         Do Until .EOF
            iRow = iRow + 1
            If iRow > 1 Then Grd.Rows = Grd.Rows + 1
'            Grd.Row = iRow
'            Grd.Col = 0
'            Grd.Text = Format(!BIDRTEOPNO, "000")
            Grd.TextMatrix(iRow, LABORCOL_OpNo) = Format(!BIDRTEOPNO, "000")
            
'            Grd.Col = 1
'            Grd.Text = "" & Trim(!BIDRTESHOP)
'            Grd.Text = GetShop(Grd.Text)
'            Grd.TextMatrix(iRow, LABORCOL_Shop) = GetShop(Grd.Text)
            Grd.TextMatrix(iRow, LABORCOL_Shop) = Trim(!BIDRTESHOP)

'            Grd.Col = 2
'            Grd.Text = "" & Trim(!BIDRTECENTER)
'            Grd.Text = GetCenter(Grd.Text)
'            Grd.TextMatrix(iRow, LABORCOL_WC) = GetCenter(Grd.Text)
            Grd.TextMatrix(iRow, LABORCOL_WC) = Trim(!BIDRTECENTER)

'            Grd.Col = 3
'            Grd.Text = Format((!BIDRTESETUP + !BIDRTEUNIT), ES_QuantityDataFormat)
            Grd.TextMatrix(iRow, LABORCOL_Hours) = Format((!BIDRTESETUP + !BIDRTEUNIT), ES_QuantityDataFormat)
            .MoveNext
         Loop
         DoEvents
         ClearResultSet RdoOps
         Grd.Row = 1
         iGridIndex = 1
         Grd.Col = 0
      End With
      If Grd.Rows > 1 Then
         txtOpn = Grd.TextMatrix(Grd.Row, LABORCOL_OpNo)
         bGoodBid = GetThisOperation()
         cmdDel.Enabled = True
      Else
         Grd.Row = 1
         txtOpn = "010"
         AddANewOperation
         cmdDel.Enabled = True
      End If
   Else
      AddANewOperation
   End If
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "getoperations"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Function GetThisOperation() As Byte
   On Error GoTo DiaErr1
   sSql = "SELECT * FROM EsrtTable WHERE BIDRTEREF=" _
          & Val(lblBid) & " AND BIDRTEOPNO=" & Val(txtOpn) & " "
   Debug.Print sSql
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoOpn, ES_KEYSET)
   If bSqlRows Then
      MouseCursor 13
      With RdoOpn
         bCanceled = 1
         'If cmbShp <> "" & Trim(!BIDRTESHOP) Then
            cmbShp = "" & Trim(!BIDRTESHOP)
            cmbShp = GetShop(cmbShp)
            sOldShop = cmbShp
            FillWorkCenters Trim(!BIDRTECENTER)
            'cmbWcn = "" & Trim(!BIDRTECENTER)
            cmbWcn = GetCenter(cmbWcn)
            sOldCenter = cmbWcn
         'End If
         txtSet = Format(!BIDRTESETUP, ES_QuantityDataFormat)
         txtUnt = Format(!BIDRTEUNIT, ES_TimeFormat)
         txtQdy = Format(!BIDRTEQHRS, ES_QuantityDataFormat)
         txtMdy = Format(!BIDRTEMHRS, ES_QuantityDataFormat)
         txtRte = Format(!BIDRTERATE, ES_QuantityDataFormat)
Debug.Print "    rate = " & !BIDRTERATE
         lblFoh = Format(!BIDRTEFOHRATE * 100, ES_QuantityDataFormat)
         Grd.Col = 3
         Grd.Text = Format(Val(txtSet) + Val(txtUnt), ES_QuantityDataFormat)
         txtCmt = "" & Trim(!BIDRTENOTES)
      End With
      'FillWorkCenters
      sOldShop = cmbShp
      Grd.Col = 0
      lblEdit = "Editing Operation " & txtOpn
      GetThisOperation = 1
      bCanceled = 0
      'On Error Resume Next
      'cmbShp.SetFocus
   Else
      sOldShop = ""
      lblEdit = "No Current Operation"
      txtSet = "0.000"
      txtQdy = "0.000"
      txtMdy = "0.000"
      txtCmt = ""
      GetThisOperation = 0
   End If
   MouseCursor 0
   Exit Function
   
DiaErr1:
   sProcName = "getthisopera"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub AddANewOperation()
   GetNextOperation
   'On Error Resume Next
'   sSql = "INSERT INTO EsrtTable (BIDRTEREF,BIDRTEOPNO," _
'          & "BIDRTESHOP,BIDRTECENTER) VALUES(" _
'          & Val(lblBid) & "," & Val(txtOpn) & ",'" _
'          & Compress(cmbShp) & "','" & Compress(cmbWcn) _
'          & "')"
'   clsADOCon.ExecuteSQL sSql ' rdExecDirect
   
   If Grd.Row > 0 Then
'      sSql = "INSERT INTO EsrtTable (BIDRTEREF,BIDRTEOPNO," & vbCrLf _
'          & "BIDRTESHOP,BIDRTECENTER,BIDRTESETUP,BIDRTEUNIT)" & vbCrLf _
'          & "VALUES(" & Val(lblBid) & "," & Val(txtOpn) & "," & vbCrLf _
'          & "'" & Compress(cmbShp) & "','" & Compress(cmbWcn) & "'," & vbCrLf _
'          & txtSet & "," & Me.txtUnt & ")"
      sSql = "INSERT INTO EsrtTable (BIDRTEREF,BIDRTEOPNO," & vbCrLf _
          & "BIDRTESHOP,BIDRTECENTER,BIDRTERATE)" & vbCrLf _
          & "VALUES(" & Val(lblBid) & "," & Val(txtOpn) & "," & vbCrLf _
          & "'" & Compress(cmbShp) & "','" & Compress(cmbWcn) & "'," & vbCrLf _
          & Me.txtRte & ")"
      clsADOCon.ExecuteSQL sSql ' rdExecDirect
   
   'If Err = 0 Then
      If Val(txtOpn) > 10 Then
         Grd.Rows = Grd.Rows + 1
         iGridIndex = Grd.Rows - 1
      Else
         iGridIndex = 1
      End If
      Grd.Row = iGridIndex
      Grd.Col = 0
      Grd.Text = txtOpn
      Grd.Col = 1
      Grd.Text = cmbShp
      Grd.Col = 2
      Grd.Text = cmbWcn
      Grd.Col = 0
      SysMsg "Operation Added.", True
      'On Error Resume Next
      bGoodBid = GetThisOperation()
      lblEdit = "Editing Operation " & txtOpn
   Else
      MsgBox "Could Not Add That Operation.", _
         vbExclamation, Caption
   End If
   
End Sub

Private Sub GetNextOperation()
   Dim RdoNxt As ADODB.Recordset
   Dim lOpNo As Integer
   sSql = "SELECT MAX(BIDRTEOPNO) FROM EsrtTable WHERE " _
          & "BIDRTEREF=" & Val(lblBid) & " "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoNxt, ES_FORWARD)
   If bSqlRows Then
      With RdoNxt
         If Not IsNull(.Fields(0)) Then
            lOpNo = .Fields(0)
         Else
            lOpNo = 0
         End If
         ClearResultSet RdoNxt
      End With
   Else
      lOpNo = 0
   End If
   txtOpn = Format(lOpNo + 10, "000")
   Set RdoNxt = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getnextopera"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Function GetShop(sShop As String) As String
   Dim RdoShp As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT SHPREF,SHPNUM FROM ShopTable " _
          & "WHERE SHPREF='" & sShop & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoShp, ES_FORWARD)
   If bSqlRows Then
      With RdoShp
         GetShop = "" & Trim(!SHPNUM)
         ClearResultSet RdoShp
      End With
   End If
   Set RdoShp = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getshop"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Function GetCenter(sCenter As String) As String
   Dim RdoWcn As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT WCNREF,WCNNUM FROM WcntTable " _
          & "WHERE WCNREF='" & sCenter & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoWcn, ES_FORWARD)
   If bSqlRows Then
      With RdoWcn
         GetCenter = "" & Trim(!WCNNUM)
         ClearResultSet RdoWcn
      End With
   End If
   Set RdoWcn = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getcenter"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

'Reorder the ops

Private Sub RewindOps()
   Dim iNew As Integer
   Dim iOPNO As Integer
   Dim iLast As Integer
   Dim RdoOps As ADODB.Recordset
   
   '1: Get current and reset them
   'On Error Resume Next
   iLast = GetLastOperation()
   sSql = "SELECT BIDRTEREF,BIDRTEOPNO FROM EsrtTable " _
          & "WHERE BIDRTEREF=" & Val(lblBid) & " ORDER BY BIDRTEOPNO "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoOps, ES_FORWARD)
   If bSqlRows Then
      MouseCursor 13
      cmdDel.Enabled = False
      iOPNO = 1020
      With RdoOps
         Do Until .EOF
            iOPNO = iOPNO + 1
            '11/18/03 Patch to force a bail (bug in forward cursor)
            If !BIDRTEOPNO > iLast Then Exit Do
            sSql = "UPDATE EsrtTable SET BIDRTEOPNO=" & iOPNO _
                   & " WHERE BIDRTEREF=" & Val(lblBid) & " " _
                   & "AND BIDRTEOPNO=" & !BIDRTEOPNO & " "
            clsADOCon.ExecuteSQL sSql ' rdExecDirect
            .MoveNext
         Loop
         ClearResultSet RdoOps
      End With
   End If
   
   '2: Get new and reset them
   If iOPNO > 1020 Then
      sSql = "SELECT BIDRTEREF,BIDRTEOPNO FROM EsrtTable " _
             & "WHERE BIDRTEREF=" & Val(lblBid) & " ORDER BY BIDRTEOPNO "
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoOps, ES_FORWARD)
      If bSqlRows Then
         iOPNO = 1020
         iNew = 0
         With RdoOps
            Do Until .EOF
               iNew = iNew + 10
               iOPNO = iOPNO + 1
               sSql = "UPDATE EsrtTable SET BIDRTEOPNO=" & iNew _
                      & " WHERE BIDRTEREF=" & Val(lblBid) & " " _
                      & "AND BIDRTEOPNO=" & iOPNO & " "
               clsADOCon.ExecuteSQL sSql ' rdExecDirect
               .MoveNext
            Loop
            ClearResultSet RdoOps
         End With
      End If
   End If
   GetOperations
   
End Sub

Private Function GetLastOperation() As Integer
   Dim RdoLst As ADODB.Recordset
   sSql = "SELECT MAX(BIDRTEOPNO) FROM EsrtTable WHERE " _
          & "BIDRTEREF=" & Val(lblBid) & " "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoLst, ES_FORWARD)
   If bSqlRows Then
      With RdoLst
         If Not IsNull(.Fields(0)) Then
            GetLastOperation = .Fields(0)
         Else
            GetLastOperation = 0
         End If
         ClearResultSet RdoLst
      End With
   Else
      GetLastOperation = 0
   End If
   Set RdoLst = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getlastopera"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub txtUnt_Validate(Cancel As Boolean)
   txtUnt = CheckLen(txtUnt, 8)
   txtUnt = Format(Abs(Val(txtUnt)), ES_TimeFormat)
   If bCanceled = 1 Then Exit Sub
   If bGoodBid = 1 Then
      'On Error Resume Next
      'RdoOpn.Edit
      RdoOpn!BIDRTEUNIT = Val(txtUnt)
      RdoOpn!BIDRTERATE = Val(txtRte)
      RdoOpn!BIDRTEFOHRATE = Val(lblFoh) / 100
      RdoOpn.Update
   End If
'   Grd.Col = 3
'   Grd.Text = Format(Val(txtSet) + Val(txtUnt), ES_QuantityDataFormat)
'   Grd.Col = 0
   Grd.TextMatrix(Grd.Row, LABORCOL_Hours) = Format(Val(txtSet) + Val(txtUnt), ES_QuantityDataFormat)
End Sub
