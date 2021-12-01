VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form ppiESe02b 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Estimate Routing "
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6240
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "ppiESe02b.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   37
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin MSFlexGridLib.MSFlexGrid Grd 
      Height          =   1500
      Left            =   360
      TabIndex        =   36
      Top             =   24
      Width           =   4452
      _ExtentX        =   7858
      _ExtentY        =   2646
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      ScrollBars      =   2
   End
   Begin VB.ComboBox txtDsc 
      Height          =   288
      Left            =   1560
      Sorted          =   -1  'True
      TabIndex        =   5
      Tag             =   "2"
      ToolTipText     =   "40 Characters Max.  Free Form (Does Not Require A Formula)"
      Top             =   2640
      Width           =   4092
   End
   Begin VB.CommandButton cmdComments 
      DisabledPicture =   "ppiESe02b.frx":07AE
      DownPicture     =   "ppiESe02b.frx":1120
      Height          =   350
      Index           =   1
      Left            =   5520
      Picture         =   "ppiESe02b.frx":1A92
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "Standard Comments"
      Top             =   3240
      Width           =   350
   End
   Begin VB.CommandButton cmdFormula 
      DisabledPicture =   "ppiESe02b.frx":2094
      DownPicture     =   "ppiESe02b.frx":2626
      Height          =   280
      Left            =   3480
      Picture         =   "ppiESe02b.frx":2BB8
      Style           =   1  'Graphical
      TabIndex        =   34
      TabStop         =   0   'False
      ToolTipText     =   "Calculate Cost With Formula"
      Top             =   2280
      Width           =   300
   End
   Begin VB.TextBox txtCost 
      Height          =   288
      Left            =   5280
      TabIndex        =   4
      Tag             =   "1"
      ToolTipText     =   "Operation Cost"
      Top             =   2280
      Width           =   852
   End
   Begin VB.ComboBox cmbFrm 
      Height          =   288
      Left            =   1560
      TabIndex        =   3
      Tag             =   "3"
      ToolTipText     =   "Formula Name (12) Characters Max - Blank For None"
      Top             =   2280
      Width           =   1815
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "&Delete"
      Enabled         =   0   'False
      Height          =   285
      Left            =   5280
      TabIndex        =   30
      TabStop         =   0   'False
      ToolTipText     =   "Delete The Current Entry"
      Top             =   600
      Width           =   885
   End
   Begin VB.TextBox txtRte 
      Height          =   285
      Left            =   4680
      TabIndex        =   8
      Tag             =   "1"
      ToolTipText     =   "Labor Rate"
      Top             =   5160
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.CommandButton cmdUpd 
      Caption         =   "&Update"
      Height          =   285
      Left            =   360
      TabIndex        =   22
      TabStop         =   0   'False
      ToolTipText     =   "Update Current Operation"
      Top             =   4800
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&Add"
      Height          =   285
      Left            =   5280
      TabIndex        =   21
      TabStop         =   0   'False
      ToolTipText     =   "Add A New Operation"
      Top             =   1800
      Width           =   885
   End
   Begin VB.TextBox txtCmt 
      Height          =   765
      Left            =   1560
      MultiLine       =   -1  'True
      TabIndex        =   11
      Tag             =   "9"
      ToolTipText     =   "Notes: 255 Char Max"
      Top             =   3240
      Width           =   3828
   End
   Begin VB.TextBox txtQdy 
      Height          =   285
      Left            =   1320
      TabIndex        =   9
      Tag             =   "1"
      ToolTipText     =   "Queue Hours"
      Top             =   4800
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.TextBox txtMdy 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   3120
      TabIndex        =   10
      Tag             =   "1"
      ToolTipText     =   "Move Hours"
      Top             =   4800
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.TextBox txtSet 
      Height          =   285
      Left            =   1320
      TabIndex        =   6
      Tag             =   "1"
      ToolTipText     =   "Lot Setup Hours"
      Top             =   5160
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.TextBox txtUnt 
      Height          =   285
      Left            =   3120
      TabIndex        =   7
      Tag             =   "1"
      ToolTipText     =   "Unit Hours"
      Top             =   5160
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.TextBox txtOpn 
      ForeColor       =   &H00400000&
      Height          =   315
      Left            =   360
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Tag             =   "1"
      ToolTipText     =   "Operation Number (Locked)"
      Top             =   1800
      Width           =   435
   End
   Begin VB.ComboBox cmbShp 
      Height          =   288
      Left            =   960
      Sorted          =   -1  'True
      TabIndex        =   1
      Tag             =   "3"
      Top             =   1800
      Width           =   1815
   End
   Begin VB.ComboBox cmbWcn 
      Height          =   288
      Left            =   2880
      TabIndex        =   2
      Tag             =   "3"
      Top             =   1800
      Width           =   1815
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5280
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5760
      Top             =   5160
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4710
      FormDesignWidth =   6240
   End
   Begin VB.Label Fr 
      BackStyle       =   0  'Transparent
      Caption         =   "Operation Cost"
      ForeColor       =   &H00400000&
      Height          =   252
      Index           =   1
      Left            =   4080
      TabIndex        =   33
      Tag             =   "2"
      ToolTipText     =   "Entered Or From Formula"
      Top             =   2280
      Width           =   1332
   End
   Begin VB.Label Fr 
      BackStyle       =   0  'Transparent
      Caption         =   "Formula"
      ForeColor       =   &H00400000&
      Height          =   252
      Index           =   2
      Left            =   360
      TabIndex        =   32
      Top             =   2280
      Width           =   1332
   End
   Begin VB.Label Fr 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      ForeColor       =   &H00400000&
      Height          =   252
      Index           =   0
      Left            =   360
      TabIndex        =   31
      ToolTipText     =   "40 Characters Max.  Free Form (Does Not Require A Formula)"
      Top             =   2640
      Width           =   1332
   End
   Begin VB.Label z1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      Height          =   288
      Index           =   10
      Left            =   5520
      TabIndex        =   29
      Top             =   4800
      Visible         =   0   'False
      Width           =   408
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "FOH"
      Height          =   288
      Index           =   7
      Left            =   4080
      TabIndex        =   28
      Top             =   4800
      Visible         =   0   'False
      Width           =   648
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Rate"
      Height          =   288
      Index           =   4
      Left            =   4080
      TabIndex        =   27
      Top             =   5160
      Visible         =   0   'False
      Width           =   648
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
      Height          =   252
      Left            =   480
      TabIndex        =   26
      Top             =   4200
      Width           =   4452
   End
   Begin VB.Label lblFoh 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   4680
      TabIndex        =   25
      ToolTipText     =   "Factory Overhead Rate"
      Top             =   4800
      Visible         =   0   'False
      Width           =   828
   End
   Begin VB.Label lblRate 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   5160
      TabIndex        =   24
      ToolTipText     =   "Labor Rate"
      Top             =   4080
      Visible         =   0   'False
      Width           =   852
   End
   Begin VB.Label lblBid 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   360
      TabIndex        =   23
      ToolTipText     =   "Total Services"
      Top             =   3240
      Visible         =   0   'False
      Width           =   852
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
      Height          =   288
      Index           =   1
      Left            =   360
      TabIndex        =   20
      Top             =   1560
      Width           =   672
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Shop                                 "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Index           =   2
      Left            =   960
      TabIndex        =   19
      Top             =   1560
      Width           =   2052
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
      Height          =   288
      Index           =   3
      Left            =   2880
      TabIndex        =   18
      Top             =   1560
      Width           =   1932
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Notes:"
      Height          =   288
      Index           =   0
      Left            =   360
      TabIndex        =   17
      Top             =   3000
      Width           =   792
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Queue Hrs"
      Height          =   288
      Index           =   5
      Left            =   360
      TabIndex        =   16
      Top             =   4800
      Visible         =   0   'False
      Width           =   1032
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Move Hrs"
      Height          =   288
      Index           =   6
      Left            =   2280
      TabIndex        =   15
      Top             =   4800
      Visible         =   0   'False
      Width           =   888
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Setup"
      Height          =   288
      Index           =   8
      Left            =   360
      TabIndex        =   14
      Top             =   5160
      Visible         =   0   'False
      Width           =   1068
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Unit"
      Height          =   288
      Index           =   9
      Left            =   2280
      TabIndex        =   13
      Top             =   5160
      Visible         =   0   'False
      Width           =   648
   End
End
Attribute VB_Name = "ppiESe02b"
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
'3/31/06 Added Comments Selection. Added Combo Box for Descriptions
'5/19/06 Reordered and fixed improper Shop/Center indexing
'5/22/06 Special Comments Index for Ron
'5/23/06 Fixed formula problem
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
Dim sOldFormula As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

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
   Set RdoPar = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getrates"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub FillWorkCenters()
   cmbWcn.Clear
   On Error GoTo DiaErr1
   sSql = "Qry_FillWorkCenters '" & Compress(cmbShp) & "'"
   LoadComboBox cmbWcn
   If cmbWcn.ListCount > 0 Then
      If sOldShop = cmbShp Then
         cmbWcn = sOldCenter
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
   FillFormulae
   Exit Sub
   
DiaErr1:
   sProcName = "fillworkc"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub cmbFrm_LostFocus()
   Dim bByte As Byte
   Dim iList As Integer
   If bCanceled = 1 Then Exit Sub
   If Trim(cmbFrm) = "" Then cmbFrm = "NONE"
   If cmbFrm <> "NONE" Then
      For iList = 0 To cmbFrm.ListCount - 1
         If cmbFrm = cmbFrm.List(iList) Then bByte = 1
      Next
      If bByte = 0 Then
         Beep
         cmbFrm = sOldFormula
      End If
   End If
   If bGoodBid = 1 Then
      On Error Resume Next
      'RdoOpn.Edit
      RdoOpn!BIDFORMULA = Trim(cmbFrm)
      RdoOpn.Update
   End If
   sOldFormula = cmbFrm
   
End Sub


Private Sub cmbShp_Click()
   GetShopInfo
   Grd.Col = 1
   If Grd.Row > 0 Then Grd.Text = cmbShp
   If sOldShop <> cmbShp Then FillWorkCenters
   
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
   FillWorkCenters
   sOldShop = cmbShp
   Grd.Col = 1
   Grd.Text = cmbShp
   If bGoodBid = 1 Then
      On Error Resume Next
      'RdoOpn.Edit
      RdoOpn!BIDRTESHOP = Compress(cmbShp)
      RdoOpn.Update
   End If
   
End Sub


Private Sub cmbWcn_Click()
   GetCenterInfo
   FillFormulae
   Grd.Col = 2
   If Grd.Row > 0 Then Grd.Text = cmbWcn
   
End Sub

Private Sub cmbWcn_LostFocus()
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
   sOldCenter = cmbWcn
   'Grd.Row = iGridIndex
   Grd.Col = 2
   Grd.Text = cmbWcn
   If bGoodBid = 1 Then
      On Error Resume Next
      'RdoOpn.Edit
      RdoOpn!BIDRTECENTER = Compress(cmbWcn)
      RdoOpn.Update
   End If
   
   
End Sub


Private Sub cmdCan_Click()
   GetTheLabor
   Unload Me
   
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   bCanceled = 1
   
End Sub


'5/22/06 Special Index for Ron

Private Sub cmdComments_Click(Index As Integer)
   Dim bIndex As Byte
   txtCmt.SetFocus
   'See List For Index
   bIndex = GetSetting("Esi2000", "ProPla", "EstimatingIndex", Trim(bIndex))
   If bIndex = 0 Then bIndex = 7
   SysComments.lblListIndex = bIndex
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
      On Error Resume Next
      RdoOpn.Close
      sSql = "DELETE FROM EsrtTable WHERE BIDRTEREF=" _
             & Val(lblBid) & " AND BIDRTEOPNO=" _
             & Val(txtOpn) & " "
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
      If Err = 0 Then
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

Private Sub cmdFormula_Click()
   ppiESe02f.cmbFrm = cmbFrm
   ppiESe02f.Show
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 8514
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
      FillDescription
      iGridIndex = 0
      Caption = Caption & " - Estimate " & lblBid
      GetRates
      lBidNo = Val(lblBid)
      FillCombo
      bOnLoad = 0
      If cmbWcn.ListCount = 0 Or cmbShp.ListCount = 0 Then
         MsgBox "There Are Either No Shops Or No Work Centers.", _
            vbExclamation, Caption
         Unload Me
      Else
         GetOperations
      End If
   End If
   MouseCursor 0
   
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
      .ColWidth(1) = 750
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


'5/25/06 Set default for BIDFORMULA

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   On Error Resume Next
   RdoOpn.Close
   sSql = "UPDATE EsrtTable SET BIDFORMULA='NONE' WHERE (BIDFORMULA='' " _
          & "AND BIDRTEREF>" & Val(lblBid) - 5 & ") "
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   Set RdoOpn = Nothing
   Set ppiESe02b = Nothing
   
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
      FillWorkCenters
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Private Sub grd_Click()
   Grd.Col = 0
   iGridIndex = Grd.Row
   txtOpn = Grd.Text
   If Val(txtOpn) > 0 Then bGoodBid = GetThisOperation()
   
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
   bCanceled = 1
   
End Sub

Private Sub txtCmt_LostFocus()
   txtCmt = CheckLen(txtCmt, 255)
   If bCanceled = 1 Then Exit Sub
   txtCmt = StrCase(txtCmt, ES_FIRSTWORD)
   If bGoodBid = 1 Then
      On Error Resume Next
      'RdoOpn.Edit
      RdoOpn!BIDRTENOTES = txtCmt
      RdoOpn.Update
   End If
   
End Sub


Private Sub txtCost_LostFocus()
   txtCost = Format(Abs(Val(txtCost)), "#####0.00")
   If bGoodBid = 1 Then
      On Error Resume Next
      'RdoOpn.Edit
      RdoOpn!BIDRTELABOR = Val(txtCost)
      RdoOpn!BIDFORMULA = Trim(cmbFrm)
      RdoOpn!BIDFORMULANOTES = Trim(txtDsc)
      RdoOpn.Update
   End If
   
End Sub


Private Sub txtDsc_LostFocus()
   Dim bByte As Byte
   Dim iList As Integer
   txtDsc = CheckLen(txtDsc, 40)
   txtDsc = StrCase(txtDsc)
   If bGoodBid = 1 Then
      On Error Resume Next
      'RdoOpn.Edit
      RdoOpn!BIDFORMULANOTES = Trim(txtDsc)
      RdoOpn.Update
   End If
   If Len(txtDsc) > 0 Then
      For iList = 0 To txtDsc.ListCount - 1
         If Trim(txtDsc) = txtDsc.List(iList) Then bByte = 1
      Next
   Else
      bByte = 1
   End If
   If bByte = 0 Then txtDsc.AddItem txtDsc
   
End Sub


Private Sub txtMdy_LostFocus()
   txtMdy = CheckLen(txtMdy, 7)
   txtMdy = Format(Abs(Val(txtMdy)), ES_QuantityDataFormat)
   If bCanceled = 1 Then Exit Sub
   If bGoodBid = 1 Then
      On Error Resume Next
      'RdoOpn.Edit
      RdoOpn!BIDRTEMHRS = Val(txtMdy)
      RdoOpn.Update
   End If
   
End Sub




Private Sub txtQdy_LostFocus()
   txtQdy = CheckLen(txtQdy, 7)
   txtQdy = Format(Abs(Val(txtQdy)), ES_QuantityDataFormat)
   If bCanceled = 1 Then Exit Sub
   If bGoodBid = 1 Then
      On Error Resume Next
      'RdoOpn.Edit
      RdoOpn!BIDRTEQHRS = Val(txtQdy)
      RdoOpn.Update
   End If
   
End Sub


Private Sub txtRte_LostFocus()
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
   Grd.Col = 3
   Grd.Text = Format(Val(txtSet) + Val(txtUnt), ES_QuantityDataFormat)
   Grd.Col = 0
   
End Sub


Private Sub txtSet_LostFocus()
   txtSet = CheckLen(txtSet, 7)
   txtSet = Format(Abs(Val(txtSet)), ES_QuantityDataFormat)
   If bCanceled = 1 Then Exit Sub
   If bGoodBid = 1 Then
      On Error Resume Next
      'RdoOpn.Edit
      RdoOpn!BIDRTESETUP = Val(txtSet)
      RdoOpn!BIDRTERATE = Val(txtRte)
      RdoOpn!BIDRTEFOHRATE = Val(lblFoh) / 100
      RdoOpn.Update
   End If
   Grd.Col = 3
   Grd.Text = Format(Val(txtSet) + Val(txtUnt), ES_QuantityDataFormat)
   Grd.Col = 0
   
End Sub


Private Sub txtUnt_LostFocus()
   txtUnt = CheckLen(txtUnt, 8)
   txtUnt = Format(Abs(Val(txtUnt)), ES_TimeFormat)
   If bCanceled = 1 Then Exit Sub
   If bGoodBid = 1 Then
      On Error Resume Next
      'RdoOpn.Edit
      RdoOpn!BIDRTEUNIT = Val(txtUnt)
      RdoOpn!BIDRTERATE = Val(txtRte)
      RdoOpn!BIDRTEFOHRATE = Val(lblFoh) / 100
      RdoOpn.Update
   End If
   Grd.Col = 3
   Grd.Text = Format(Val(txtSet) + Val(txtUnt), ES_QuantityDataFormat)
   Grd.Col = 0
   
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
   If cCurrentRate = 0 Then cCurrentRate = cDefaultRate
   If sOldCenter <> cmbWcn Or Val(txtRte) = 0 _
                                  Then txtRte = Format(cCurrentRate, ES_QuantityDataFormat)
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
               Grd.Row = iRow
               Grd.Col = 0
               Grd.Text = Format(!BIDRTEOPNO, "000")
               Grd.Col = 1
               Grd.Text = "" & Trim(!BIDRTESHOP)
               Grd.Text = GetShop(Grd.Text)
               Grd.Col = 2
               Grd.Text = "" & Trim(!BIDRTECENTER)
               Grd.Text = GetCenter(Grd.Text)
               Grd.Col = 3
               Grd.Text = Format((!BIDRTESETUP + !BIDRTEUNIT), ES_QuantityDataFormat)
               .MoveNext
            Loop
            ClearResultSet RdoOps
            Grd.Row = 1
            iGridIndex = 1
            Grd.Col = 0
         End With
         If Grd.Rows > 1 Then
            txtOpn = Grd.Text
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
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoOpn, ES_KEYSET)
      If bSqlRows Then
         MouseCursor 13
         With RdoOpn
            bCanceled = 1
            cmbShp = "" & Trim(!BIDRTESHOP)
            cmbShp = GetShop(cmbShp)
            sOldShop = cmbShp
            cmbWcn = "" & Trim(!BIDRTECENTER)
            cmbWcn = GetCenter(cmbWcn)
            sOldCenter = cmbWcn
            cmbFrm = "" & Trim(!BIDFORMULA)
            sOldFormula = cmbFrm
            txtDsc = "" & Trim(!BIDFORMULANOTES)
            txtCost = Format(!BIDRTELABOR, "#####0.00")
            sOldFormula = cmbFrm
            Grd.Col = 3
            Grd.Text = Format(Val(txtSet) + Val(txtUnt), ES_QuantityDataFormat)
            txtCmt = "" & Trim(!BIDRTENOTES)
         End With
         FillWorkCenters
         'FillFormulae
         Grd.Col = 0
         lblEdit = "Editing Operation " & txtOpn
         GetThisOperation = 1
         bCanceled = 0
         On Error Resume Next
         cmbShp.SetFocus
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
      On Error Resume Next
      sSql = "INSERT INTO EsrtTable (BIDRTEREF,BIDRTEOPNO," _
             & "BIDRTESHOP,BIDRTECENTER) VALUES(" _
             & Val(lblBid) & "," & Val(txtOpn) & ",'" _
             & Compress(cmbShp) & "','" & Compress(cmbWcn) _
             & "')"
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
      If Err = 0 Then
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
         On Error Resume Next
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
      On Error Resume Next
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
               clsADOCon.ExecuteSQL sSql 'rdExecDirect
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
                  clsADOCon.ExecuteSQL sSql 'rdExecDirect
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
   
   Private Sub GetTheLabor()
      Dim RdoLabor As ADODB.Recordset
      On Error Resume Next
      sSql = "SELECT SUM(BIDRTELABOR) AS BidLabor FROM EsrtTable WHERE BIDRTEREF=" _
             & Val(lblBid)
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoLabor, ES_FORWARD)
      If bSqlRows Then
         With RdoLabor
            ppiESe02a.lblLabor = Format(!Bidlabor, "####0.00")
            .Cancel
         End With
      Else
         ppiESe02a.lblLabor = "0.00"
      End If
      
   End Sub
   
   Private Sub FillFormulae()
      cmbFrm.Clear
      sSql = "SELECT FORMULA_REF FROM EsfrTable WHERE (FORMULA_REF<>'NONE' " _
             & "AND FORMULA_CENTER='" & Compress(cmbWcn) & "') ORDER BY FORMULA_REF"
      LoadComboBox cmbFrm, -1
      cmbFrm.AddItem "NONE"
      If cmbFrm.ListCount > 0 Then
         cmdFormula.Enabled = True
         cmbFrm = sOldFormula
      Else
         cmdFormula.Enabled = False
         cmbFrm = "NONE"
      End If
      
   End Sub
   
   Private Sub FillDescription()
      sSql = "SELECT DISTINCT BIDFORMULANOTES FROM EsrtTable " _
             & "WHERE BIDFORMULANOTES<>''"
      LoadComboBox txtDsc, -1
      
   End Sub
