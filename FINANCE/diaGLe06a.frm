VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaGLe06a 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Account Budgets"
   ClientHeight    =   7140
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5265
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7140
   ScaleWidth      =   5265
   Tag             =   "1"
   Begin VB.TextBox txtAmt 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Index           =   12
      Left            =   2880
      TabIndex        =   15
      Tag             =   "1"
      Top             =   6660
      Width           =   1500
   End
   Begin VB.TextBox txtAmt 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Index           =   11
      Left            =   2880
      TabIndex        =   14
      Tag             =   "1"
      Top             =   6300
      Width           =   1500
   End
   Begin VB.TextBox txtAmt 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Index           =   10
      Left            =   2880
      TabIndex        =   13
      Tag             =   "1"
      Top             =   5940
      Width           =   1500
   End
   Begin VB.TextBox txtAmt 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Index           =   9
      Left            =   2880
      TabIndex        =   12
      Tag             =   "1"
      Top             =   5580
      Width           =   1500
   End
   Begin VB.TextBox txtAmt 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Index           =   8
      Left            =   2880
      TabIndex        =   11
      Tag             =   "1"
      Top             =   5220
      Width           =   1500
   End
   Begin VB.TextBox txtAmt 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Index           =   7
      Left            =   2880
      TabIndex        =   10
      Tag             =   "1"
      Top             =   4860
      Width           =   1500
   End
   Begin VB.TextBox txtAmt 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Index           =   6
      Left            =   2880
      TabIndex        =   9
      Tag             =   "1"
      Top             =   4500
      Width           =   1500
   End
   Begin VB.TextBox txtAmt 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Index           =   5
      Left            =   2880
      TabIndex        =   8
      Tag             =   "1"
      Top             =   4140
      Width           =   1500
   End
   Begin VB.TextBox txtAmt 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Index           =   4
      Left            =   2880
      TabIndex        =   7
      Tag             =   "1"
      Top             =   3780
      Width           =   1500
   End
   Begin VB.TextBox txtAmt 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Index           =   3
      Left            =   2880
      TabIndex        =   6
      Tag             =   "1"
      Top             =   3420
      Width           =   1500
   End
   Begin VB.TextBox txtAmt 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   2880
      TabIndex        =   5
      Tag             =   "1"
      Top             =   3060
      Width           =   1500
   End
   Begin VB.TextBox txtAmt 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   2880
      TabIndex        =   4
      Tag             =   "1"
      Top             =   2700
      Width           =   1500
   End
   Begin VB.TextBox txtAmt 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   2880
      TabIndex        =   3
      Tag             =   "1"
      Top             =   2340
      Width           =   1500
   End
   Begin VB.ComboBox cmbAct 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1440
      TabIndex        =   2
      Top             =   1320
      Width           =   1935
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   4200
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   120
      Width           =   875
   End
   Begin VB.ComboBox cmbRoot 
      Height          =   315
      Left            =   1440
      TabIndex        =   1
      Top             =   720
      Width           =   1935
   End
   Begin VB.ComboBox cmbYer 
      Height          =   315
      Left            =   1440
      TabIndex        =   0
      Tag             =   "1"
      Top             =   360
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   35
      Left            =   120
      TabIndex        =   19
      Top             =   1080
      Width           =   4935
   End
   Begin VB.CommandButton cmdDsl 
      Caption         =   "&Cancel"
      Enabled         =   0   'False
      Height          =   315
      Left            =   4200
      TabIndex        =   16
      ToolTipText     =   "Select A Different Year"
      Top             =   1320
      Width           =   875
   End
   Begin VB.CommandButton cmdSel 
      Caption         =   "&Select"
      Height          =   315
      Left            =   4200
      TabIndex        =   18
      ToolTipText     =   "Update Structure And Associated Entries"
      Top             =   600
      Width           =   875
   End
   Begin VB.CommandButton cmdUpd 
      Caption         =   "&Update"
      Enabled         =   0   'False
      Height          =   315
      Left            =   4200
      TabIndex        =   17
      ToolTipText     =   "Update Budgets For This Account"
      Top             =   1680
      Width           =   875
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   21
      ToolTipText     =   "Subject Help"
      Top             =   0
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   397
      _StockProps     =   65
      BackColor       =   12632256
      Autosize        =   2
      RoundedCorners  =   0   'False
      BevelWidth      =   0
      Outline         =   0   'False
      PictureUp       =   "diaGLe06a.frx":0000
      PictureDn       =   "diaGLe06a.frx":0146
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Period"
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
      Left            =   120
      TabIndex        =   68
      Top             =   2100
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label lblEnd 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   12
      Left            =   1800
      TabIndex        =   67
      Top             =   6660
      Width           =   975
   End
   Begin VB.Label lblEnd 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   11
      Left            =   1800
      TabIndex        =   66
      Top             =   6300
      Width           =   975
   End
   Begin VB.Label lblEnd 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   10
      Left            =   1800
      TabIndex        =   65
      Top             =   5940
      Width           =   975
   End
   Begin VB.Label lblEnd 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   9
      Left            =   1800
      TabIndex        =   64
      Top             =   5580
      Width           =   975
   End
   Begin VB.Label lblEnd 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   8
      Left            =   1800
      TabIndex        =   63
      Top             =   5220
      Width           =   975
   End
   Begin VB.Label lblEnd 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   7
      Left            =   1800
      TabIndex        =   62
      Top             =   4860
      Width           =   975
   End
   Begin VB.Label lblEnd 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   6
      Left            =   1800
      TabIndex        =   61
      Top             =   4500
      Width           =   975
   End
   Begin VB.Label lblEnd 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   5
      Left            =   1800
      TabIndex        =   60
      Top             =   4140
      Width           =   975
   End
   Begin VB.Label lblEnd 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   4
      Left            =   1800
      TabIndex        =   59
      Top             =   3780
      Width           =   975
   End
   Begin VB.Label lblEnd 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   3
      Left            =   1800
      TabIndex        =   58
      Top             =   3420
      Width           =   975
   End
   Begin VB.Label lblEnd 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   2
      Left            =   1800
      TabIndex        =   57
      Top             =   3060
      Width           =   975
   End
   Begin VB.Label lblEnd 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   1
      Left            =   1800
      TabIndex        =   56
      Top             =   2700
      Width           =   975
   End
   Begin VB.Label lblStart 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   12
      Left            =   720
      TabIndex        =   55
      Top             =   6660
      Width           =   975
   End
   Begin VB.Label lblStart 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   11
      Left            =   720
      TabIndex        =   54
      Top             =   6300
      Width           =   975
   End
   Begin VB.Label lblStart 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   10
      Left            =   720
      TabIndex        =   53
      Top             =   5940
      Width           =   975
   End
   Begin VB.Label lblStart 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   9
      Left            =   720
      TabIndex        =   52
      Top             =   5580
      Width           =   975
   End
   Begin VB.Label lblStart 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   8
      Left            =   720
      TabIndex        =   51
      Top             =   5220
      Width           =   975
   End
   Begin VB.Label lblStart 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   7
      Left            =   720
      TabIndex        =   50
      Top             =   4860
      Width           =   975
   End
   Begin VB.Label lblStart 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   6
      Left            =   720
      TabIndex        =   49
      Top             =   4500
      Width           =   975
   End
   Begin VB.Label lblStart 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   5
      Left            =   720
      TabIndex        =   48
      Top             =   4140
      Width           =   975
   End
   Begin VB.Label lblStart 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   4
      Left            =   720
      TabIndex        =   47
      Top             =   3780
      Width           =   975
   End
   Begin VB.Label lblStart 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   3
      Left            =   720
      TabIndex        =   46
      Top             =   3420
      Width           =   975
   End
   Begin VB.Label lblStart 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   2
      Left            =   720
      TabIndex        =   45
      Top             =   3060
      Width           =   975
   End
   Begin VB.Label lblStart 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   1
      Left            =   720
      TabIndex        =   44
      Top             =   2700
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Budget Amount        "
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
      Index           =   19
      Left            =   2880
      TabIndex        =   43
      Top             =   2100
      Width           =   1545
   End
   Begin VB.Label lblEnd 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   0
      Left            =   1800
      TabIndex        =   42
      Top             =   2340
      Width           =   975
   End
   Begin VB.Label lblStart 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   0
      Left            =   720
      TabIndex        =   41
      Top             =   2340
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ending          "
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
      Index           =   18
      Left            =   1800
      TabIndex        =   40
      Top             =   2100
      Width           =   1155
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Starting         "
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
      Index           =   17
      Left            =   720
      TabIndex        =   39
      Top             =   2100
      Width           =   1035
   End
   Begin VB.Label z1 
      Caption         =   "Account"
      Height          =   255
      Index           =   16
      Left            =   360
      TabIndex        =   38
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1440
      TabIndex        =   37
      Top             =   1680
      Width           =   2400
   End
   Begin VB.Label z1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   315
      Index           =   2
      Left            =   330
      TabIndex        =   36
      Top             =   2340
      Width           =   195
   End
   Begin VB.Label z1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      Height          =   315
      Index           =   4
      Left            =   330
      TabIndex        =   35
      Top             =   2700
      Width           =   195
   End
   Begin VB.Label z1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      Height          =   315
      Index           =   5
      Left            =   330
      TabIndex        =   34
      Top             =   3060
      Width           =   195
   End
   Begin VB.Label z1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      Height          =   315
      Index           =   6
      Left            =   330
      TabIndex        =   33
      Top             =   3420
      Width           =   195
   End
   Begin VB.Label z1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      Height          =   315
      Index           =   7
      Left            =   330
      TabIndex        =   32
      Top             =   3780
      Width           =   195
   End
   Begin VB.Label z1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      Height          =   315
      Index           =   8
      Left            =   330
      TabIndex        =   31
      Top             =   4140
      Width           =   195
   End
   Begin VB.Label z1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "7"
      Height          =   315
      Index           =   9
      Left            =   330
      TabIndex        =   30
      Top             =   4500
      Width           =   195
   End
   Begin VB.Label z1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "8"
      Height          =   315
      Index           =   10
      Left            =   330
      TabIndex        =   29
      Top             =   4860
      Width           =   195
   End
   Begin VB.Label z1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "9"
      Height          =   315
      Index           =   11
      Left            =   330
      TabIndex        =   28
      Top             =   5220
      Width           =   195
   End
   Begin VB.Label z1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      Height          =   315
      Index           =   12
      Left            =   330
      TabIndex        =   27
      Top             =   5580
      Width           =   195
   End
   Begin VB.Label z1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "11"
      Height          =   315
      Index           =   13
      Left            =   330
      TabIndex        =   26
      Top             =   5940
      Width           =   195
   End
   Begin VB.Label z1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "12"
      Height          =   315
      Index           =   14
      Left            =   330
      TabIndex        =   25
      Top             =   6300
      Width           =   195
   End
   Begin VB.Label z1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "13"
      Height          =   315
      Index           =   15
      Left            =   330
      TabIndex        =   24
      Top             =   6660
      Width           =   195
   End
   Begin VB.Label z1 
      Caption         =   "Fiscal Year"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   23
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label z1 
      Caption         =   "Root Acct."
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   22
      Top             =   720
      Width           =   975
   End
End
Attribute VB_Name = "diaGLe06a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of                     ***
'*** ESI Software Engineering, Inc, Seattle, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***

'See the UpdateTables prodecure for database revisions

'Edit Budgets (JCW) 1/30/04
'Revisions:
'   (2/27/04) (JCW) Add error handling to functions
'*************************************************************************

Option Explicit

Dim bCanceled As Byte
Dim bOnLoad As Byte
Dim sCurrentAccount As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd







'***************************** EVENTS ************************************

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   sCurrForm = Caption
   bOnLoad = True
   
End Sub


Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillCombo
      bOnLoad = False
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   On Error Resume Next
   'set resultsets = nothin
   Set diaGLe06a = Nothing
   
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      SelectHelpTopic Me, Caption
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   bCanceled = True
   
End Sub

Private Sub cmdSel_Click()
   On Error Resume Next
   
   ManageBoxes True 'Sets all boxes true or false
   LoadFiscalPeriods Val(cmbYer)
   FillDivisionAccounts cmbRoot 'Fills Account Combo Box
   cmbRoot.SelLength = 0
   cmbYer.SelLength = 0
   
End Sub

Private Sub cmbYer_LostFocus()
   If Not bValidElement(cmbYer) Then
      cmdSel.enabled = False
   Else
      cmdSel.enabled = True
   End If
   
End Sub

Private Sub cmdDsl_Click()
   On Error Resume Next
   ManageBoxes False
   cmbAct.SelLength = 0
   
End Sub

Private Sub cmdUpd_Click()
   On Error Resume Next
   
   If BudgetExists(cmbAct) Then
      UpdateBudgets (cmbAct)
   Else
      If NewBudget(cmbAct) Then
         UpdateBudgets (cmbAct)
      Else
         MsgBox "Could Not Update Budgets.", vbExclamation, Caption
      End If
   End If
   
End Sub

Private Sub cmbRoot_LostFocus()
   cmbRoot = CheckLen(cmbRoot, 12)
   
End Sub

Private Sub cmbAct_Click()
   On Error Resume Next
   If Trim(cmbAct) <> sCurrentAccount Then
      FindAccount Me
      ManageBudgetBox True
      ManageLowButtons True
      LoadBudgets (cmbAct)
      sCurrentAccount = Trim(cmbAct)
   End If
   
End Sub

Private Sub cmbAct_LostFocus()
   On Error Resume Next
   'on error resume next
   cmbAct = CheckLen(cmbAct, 12)
   If Trim(cmbAct) <> "" Then
      FindAccount Me
   Else
      lblDsc = ""
   End If
   If Left(lblDsc, 10) = "*** Accoun" Or Trim(cmbAct) = "" Then
      If cmdDsl.enabled = True Then
         cmdDsl.SetFocus
      End If
      ManageBudgetBox (False)
      ManageLowButtons (False)
      
   Else
      If Trim(cmbAct) <> sCurrentAccount Then
         ManageBudgetBox (True)
         ManageLowButtons (True)
         LoadBudgets (cmbAct)
         
         If txtAmt(0).enabled = True Then txtAmt(0).SetFocus
      End If
   End If
   sCurrentAccount = Trim(cmbAct)
   
End Sub

Private Sub txtAmt_GotFocus(Index As Integer)
   SelectFormat Me
   
End Sub

Private Sub txtAmt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub

Private Sub txtAmt_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyValue KeyAscii
   
End Sub

Private Sub txtAmt_LostFocus(Index As Integer)
   On Error Resume Next
   If Left(lblDsc, 10) <> "*** Accoun" And Trim(cmbAct) <> "" Then
      If Trim(txtAmt(Index)) = "" And txtAmt(Index).enabled = True Then txtAmt(Index) = "0"
      txtAmt(Index) = NumberFix(txtAmt(Index))
      txtAmt(Index) = Format(txtAmt(Index), CURRENCYMASK)
   End If
   
End Sub


'************************** COMBO FILLERS ********************************

Public Sub FillCombo()
   Dim rdoCombo As ADODB.Recordset
   Dim i As Integer
   On Error GoTo DiaErr1
   
   sSql = "Qry_FillAccountCombo"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoCombo)
   If bSqlRows Then
      With rdoCombo
         Do While Not .EOF
            AddComboStr cmbRoot.hWnd, "" & !GLACCTNO
            rdoCombo.MoveNext
         Loop
      End With
   End If
   
   If cmbRoot.ListCount > 0 Then
      cmbRoot.ListIndex = 0
   End If
   
   
   Set rdoCombo = Nothing
   
   sSql = "SELECT FYYEAR from GlfyTable "
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoCombo)
   If bSqlRows Then
      With rdoCombo
         Do While Not .EOF
            AddComboStr cmbYer.hWnd, "" & !FYYEAR
            rdoCombo.MoveNext
         Loop
      End With
   End If
   
   If cmbYer.ListCount > 0 Then
      For i = 0 To cmbYer.ListCount
         If cmbYer.List(i) = Format(Now, "yyyy") Then
            cmbYer.ListIndex = i
            Exit For
         End If
      Next
      If Trim(cmbYer) = "" Then cmbYer.ListIndex = 0
   End If
   Set rdoCombo = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub FillDivisionAccounts(sRoot As String)
   Dim rdoRoot As ADODB.Recordset
   On Error GoTo DiaErr1
   
   sSql = "SELECT GLACCTNO FROM GlacTable WHERE LEFT(GLACCTNO," & Len(sRoot) & ")" _
          & " = '" & sRoot & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoRoot)
   If bSqlRows Then
      With rdoRoot
         Do Until .EOF
            AddComboStr cmbAct.hWnd, "" & !GLACCTNO
            .MoveNext
         Loop
      End With
   End If
   If cmbAct.ListCount > 0 Then
      cmbAct.ListIndex = 0
   End If
   Set rdoRoot = Nothing
   
   Exit Sub
   
DiaErr1:
   sProcName = "FillDivisionAccounts"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

'****************************** DATA RETRIEVAL ***************************

Private Sub LoadFiscalPeriods(iYear As Integer)
   Dim rdoYer As ADODB.Recordset
   Dim i As Integer
   Dim X As Integer
   
   On Error GoTo DiaErr1
   
   '(Below) Not "SELECT *" So that if Table is Changed We can still call .Fields(i)
   sSql = "SELECT FYPERSTART1, FYPEREND1, FYPERSTART2, " _
          & " FYPEREND2, FYPERSTART3, FYPEREND3, FYPERSTART4," _
          & " FYPEREND4, FYPERSTART5, FYPEREND5, FYPERSTART6," _
          & " FYPEREND6, FYPERSTART7, FYPEREND7, FYPERSTART8," _
          & " FYPEREND8, FYPERSTART9, FYPEREND9, FYPERSTART10," _
          & " FYPEREND10, FYPERSTART11, FYPEREND11, FYPERSTART12," _
          & " FYPEREND12, FYPERSTART13, FYPEREND13, FYPERIODS " _
          & " From GlfyTable WHERE FYYEAR = " & iYear
   
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoYer)
   
   If bSqlRows Then
      With rdoYer
         For i = 0 To Val((!FYPERIODS * 2) - 2) Step 2
            lblStart(X) = Format(.Fields(i), "mm/dd/yy")
            lblEnd(X) = Format(.Fields(i + 1), "mm/dd/yy")
            X = X + 1
         Next
      End With
   End If
   
   Set rdoYer = Nothing
   
   Exit Sub
   
DiaErr1:
   sProcName = "LoadFiscalPeriods"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub LoadBudgets(sAccount As String)
   Dim rdoBudget As ADODB.Recordset
   Dim i As Integer
   On Error GoTo DiaErr1
   
   '(Below) Not "SELECT *" So that if Table is Changed We can still call .Fields(i)
   sSql = " SELECT BUDPER1, BUDPER2, BUDPER3, BUDPER4, BUDPER5, BUDPER6," _
          & " BUDPER7, BUDPER8, BUDPER9, BUDPER10, BUDPER11, BUDPER12," _
          & " BUDPER13 FROM BdgtTable WHERE BUDFY = " & cmbYer & " AND " _
          & " BUDACCT = '" & sAccount & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoBudget)
   
   If bSqlRows Then
      With rdoBudget
         For i = 0 To 12
            If Trim(lblStart(i)) <> "" Then
               txtAmt(i) = Format(Val(.Fields(i)), CURRENCYMASK)
            End If
         Next
      End With
   Else
      For i = 0 To 12
         If Trim(lblStart(i)) <> "" Then
            txtAmt(i) = "0.00"
         End If
      Next
   End If
   Set rdoBudget = Nothing
   
   Exit Sub
   
DiaErr1:
   sProcName = "LoadBudgets"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub


Private Sub ManageBoxes(bBool As Boolean)
   Dim i As Integer
   Dim bOpp As Boolean
   
   On Error GoTo DiaErr1
   
   'determine what bOpp is:
   'bopp = CBool(CInt(bopp) + 1)
   If bBool = True Then
      bOpp = False
   Else
      bOpp = True
   End If
   
   'bBool Actions
   ManageBudgetBox (bBool) ' Sepperate Functions Because Sometimes We want to call them sepperately
   ManageLowButtons (bBool) '   when an invalid account is entered...(disable the update and fill buttons)
   cmbAct.enabled = bBool
   cmdDsl.enabled = bBool
   
   'Clear Everything
   cmbAct.Clear
   cmbAct = ""
   lblDsc = ""
   For i = 0 To 12
      lblStart(i) = ""
      lblEnd(i) = ""
   Next
   
   'bOpp Actions
   cmbYer.enabled = bOpp
   cmbRoot.enabled = bOpp
   cmdSel.enabled = bOpp
   
   If bBool = True Then
      cmbAct.SetFocus
   Else
      cmbRoot.SetFocus
   End If
   
   Exit Sub
DiaErr1:
   sProcName = "ManageBoxes"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub ManageBudgetBox(bBool As Boolean) 'Sepperate Function (so we can single it out)
   Dim i As Integer
   On Error GoTo DiaErr1
   
   For i = 0 To 12
      If (bBool = True And Trim(lblStart(i)) <> "") Or (bBool = False) Then
         'If were activating, activate only valid boxes; if not lock down everything
         txtAmt(i).enabled = bBool
         txtAmt(i) = ""
      End If
   Next
   
   Exit Sub
   
DiaErr1:
   sProcName = "ManagebudgetBox"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub ManageLowButtons(bBool As Boolean)
   On Error Resume Next
   cmdUpd.enabled = bBool
   
End Sub



Private Function NewBudget(sAccount As String) As Byte
   On Error GoTo DiaErr1
   sSql = "INSERT INTO BdgtTable (BUDACCT,BUDFY)" _
          & " VALUES ('" & Compress(sAccount) & "'," & cmbYer & ")"
   clsADOCon.ExecuteSQL sSql
   If clsADOCon.RowsAffected Then
      NewBudget = True
   Else
      NewBudget = False
   End If
   
   Exit Function
DiaErr1:
   sProcName = "NewBudget"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub UpdateBudgets(sAccount As String)
   Dim i As Integer
   Dim nAmount As Single
   On Error Resume Next
   Err.Clear
   clsADOCon.BeginTrans
   clsADOCon.ADOErrNum = 0
   For i = 0 To 12
      If Trim(txtAmt(i)) <> "" Then
         nAmount = CCur(txtAmt(i))
      Else
         nAmount = 0
      End If
      sSql = "UPDATE BdgtTable SET BUDPER" & i + 1 & " = " & nAmount _
             & " WHERE BUDACCT = '" & Compress(sAccount) & "' AND BUDFY = " _
             & cmbYer
      clsADOCon.ExecuteSQL sSql
   Next
   
   If clsADOCon.ADOErrNum = 0 Then
      clsADOCon.CommitTrans
      SysMsg "Updating Budgets.", True
   Else
      clsADOCon.RollbackTrans
      clsADOCon.ADOErrNum = 0
      MsgBox "Unable To Update Budgets.", vbExclamation, Caption
      GoTo DiaErr1
   End If
   
   Exit Sub
   
DiaErr1:
   sProcName = "UpdateBudgets"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub



Private Function BudgetExists(sAccount As String) As Boolean
   Dim rdoExst As ADODB.Recordset
   On Error GoTo DiaErr1
   
   sSql = "SELECT BUDFY, BUDACCT FROM BdgtTable WHERE BUDACCT = '" _
          & Compress(sAccount) & "' AND BUDFY = " & cmbYer
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoExst)
   If bSqlRows Then
      BudgetExists = True
   Else
      BudgetExists = False
   End If
   Set rdoExst = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "budgetexists"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Function NumberFix(sNumber As String) As String
   Dim i As Integer
   Dim sRight As String
   
   'Fixes Commas on the far left, removes multiple decimals
   
   If Left(sNumber, 1) = "," Then
      sNumber = Right(sNumber, Len(sNumber) - 1)
   End If
   
   For i = 1 To Len(sNumber)
      If InStr(Right(sNumber, i), ".") Then
         sRight = Right(sNumber, i)
         Exit For
      End If
   Next
   i = 0
   
   sNumber = Left(sNumber, Len(sNumber) - Len(sRight))
   RemoveSymbols sNumber
   RemoveSymbols sRight
   If Trim(sNumber) <> "" Or Trim(sRight) <> "" Then
      NumberFix = sNumber & "." & sRight
   Else
      NumberFix = "0"
   End If
   Exit Function
   
DiaErr1:
   sProcName = "NumberFix"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Function

Private Sub RemoveSymbols(sNum As String)
   Dim i As Integer
   On Error GoTo DiaErr1
   For i = 1 To Len(sNum)
      If InStr(Left(sNum, i), ".") Or InStr(Left(sNum, i), "-") Or InStr(Left(sNum, i), "+") Then
         'delete decimal and return the string without
         sNum = Left(sNum, i - 1) & Right(sNum, Len(sNum) - i)
         i = i - 1
      End If
   Next
   sNum = Left(sNum, 13)
   Exit Sub
   
DiaErr1:
   sProcName = "RemoveSymbols"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Function bValidElement(cmbCombo As ComboBox) As Boolean
   Dim i As Integer
   On Error GoTo DiaErr1
   
   If cmbCombo.ListCount > 0 Then
      For i = 0 To cmbCombo.ListCount - 1
         If Val(cmbCombo.List(i)) = Val(cmbCombo.Text) Then
            bValidElement = True
            cmbCombo.ListIndex = i
         End If
      Next
   End If
   Exit Function
   
DiaErr1:
   sProcName = "ValidElement"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Function
