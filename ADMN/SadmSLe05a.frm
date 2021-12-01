VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form SadmSLe05a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "State Codes"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5040
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   HelpContextID   =   1205
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   5040
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ProgressBar prg1 
      Height          =   252
      Left            =   1800
      TabIndex        =   13
      Top             =   2160
      Width           =   2532
      _ExtentX        =   4471
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "SadmSLe05a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox cmbReg 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   3240
      Sorted          =   -1  'True
      TabIndex        =   3
      ToolTipText     =   "Select Region From List"
      Top             =   1320
      Width           =   660
   End
   Begin VB.ComboBox cmbDiv 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1440
      Sorted          =   -1  'True
      TabIndex        =   2
      ToolTipText     =   "Select Division From List"
      Top             =   1320
      Width           =   860
   End
   Begin VB.CheckBox optDef 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   1440
      TabIndex        =   4
      Top             =   1800
      Width           =   735
   End
   Begin VB.ComboBox cmbSte 
      Height          =   315
      Left            =   1440
      Sorted          =   -1  'True
      TabIndex        =   0
      ToolTipText     =   "Enter/Revise A Region (2 char)"
      Top             =   600
      Width           =   660
   End
   Begin VB.TextBox txtDsc 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Top             =   960
      Width           =   3075
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   4080
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   4560
      Top             =   1920
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2655
      FormDesignWidth =   5040
   End
   Begin VB.Label lblSte 
      BackStyle       =   0  'Transparent
      Caption         =   "Building States"
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
      Left            =   240
      TabIndex        =   11
      Top             =   2160
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Default State"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   10
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Region"
      Height          =   255
      Index           =   1
      Left            =   2520
      TabIndex        =   9
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Division"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   8
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "State Code"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   7
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   6
      Top             =   960
      Width           =   1215
   End
End
Attribute VB_Name = "SadmSLe05a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions

Dim RdoCde As ADODB.Recordset
Dim bOnLoad As Byte
Dim bGoodState As Boolean

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbDiv_LostFocus()
   On Error Resume Next
   'RdoCde.Edit
   RdoCde!STATEDIV = "" & cmbDiv
   RdoCde.Update
   If Err > 0 Then ValidateEdit
   
End Sub


Private Sub cmbReg_LostFocus()
   On Error Resume Next
   'RdoCde.Edit
   RdoCde!STATEREG = "" & cmbReg
   RdoCde.Update
   If Err > 0 Then ValidateEdit
   
End Sub


Private Sub cmbSte_Click()
   bGoodState = GetState()
   
End Sub


Private Sub cmbSte_LostFocus()
   cmbSte = CheckLen(cmbSte, 2)
   If Len(cmbSte) Then
      bGoodState = GetState()
   Else
      bGoodState = False
      Exit Sub
   End If
   If Not bGoodState Then AddCode
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   cmbSte = ""
   
End Sub


Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 1205
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub


Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      FillDivisions
      FillRegions
      FillCmbStates
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
   Set RdoCde = Nothing
   Set SadmSLe05a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub BuildStateCodes()
   Dim iList As Integer
   Dim sStates(50, 2) As String
   
   MouseCursor 13
   prg1.Visible = True
   lblSte.Visible = True
   prg1.value = 5
   lblSte.Refresh
   
   On Error GoTo DiaErr1
   sStates(0, 0) = "WA"
   sStates(0, 1) = "Washington"
   
   sStates(1, 0) = "OR"
   sStates(1, 1) = "Oregon"
   
   sStates(2, 0) = "CA"
   sStates(2, 1) = "California"
   
   sStates(3, 0) = "ID"
   sStates(3, 1) = "Idaho"
   
   sStates(4, 0) = "NV"
   sStates(4, 1) = "Nevada"
   
   sStates(5, 0) = "UT"
   sStates(5, 1) = "Utah"
   
   sStates(6, 0) = "AZ"
   sStates(6, 1) = "Arizona"
   
   sStates(7, 0) = "MT"
   sStates(7, 1) = "Montana"
   
   sStates(8, 0) = "WY"
   sStates(8, 1) = "Wyoming"
   
   sStates(9, 0) = "CO"
   sStates(9, 1) = "Colorado"
   
   sStates(10, 0) = "NM"
   sStates(10, 1) = "New Mexico"
   
   sStates(11, 0) = "ND"
   sStates(11, 1) = "north Dakota"
   
   sStates(12, 0) = "SD"
   sStates(12, 1) = "South Dakota"
   
   sStates(13, 0) = "NE"
   sStates(13, 1) = "Nebraska"
   
   sStates(14, 0) = "KS"
   sStates(14, 1) = "Kansas"
   
   sStates(15, 0) = "OK"
   sStates(15, 1) = "Oklahoma"
   
   sStates(16, 0) = "TX"
   sStates(16, 1) = "Texas"
   
   sStates(17, 0) = "MN"
   sStates(17, 1) = "Minnesota"
   
   sStates(18, 0) = "IA"
   sStates(18, 1) = "Iowa"
   
   sStates(19, 0) = "MO"
   sStates(19, 1) = "Missouri"
   
   sStates(20, 0) = "AR"
   sStates(20, 1) = "Arkansas"
   
   sStates(21, 0) = "LA"
   sStates(21, 1) = "Louisiana"
   
   sStates(22, 0) = "WI"
   sStates(22, 1) = "Wisconsin"
   
   sStates(23, 0) = "IL"
   sStates(23, 1) = "Illinois"
   
   sStates(24, 0) = "MI"
   sStates(24, 1) = "Michigan"
   
   sStates(25, 0) = "IN"
   sStates(25, 1) = "Indiana"
   
   prg1.value = 10
   sStates(26, 0) = "KY"
   sStates(26, 1) = "Kentucky"
   
   sStates(27, 0) = "TN"
   sStates(27, 1) = "Tennessee"
   
   sStates(28, 0) = "MS"
   sStates(28, 1) = "Mississippi"
   
   sStates(29, 0) = "AL"
   sStates(29, 1) = "Alabama"
   
   sStates(30, 0) = "GA"
   sStates(30, 1) = "Georgia"
   
   sStates(31, 0) = "FL"
   sStates(31, 1) = "Florida"
   
   sStates(32, 0) = "ME"
   sStates(32, 1) = "Maine"
   
   sStates(33, 0) = "NH"
   sStates(33, 1) = "New Hampshire"
   
   sStates(34, 0) = "VT"
   sStates(34, 1) = "Vermont"
   
   sStates(35, 0) = "MA"
   sStates(35, 1) = "Massachusetts"
   
   sStates(36, 0) = "RI"
   sStates(36, 1) = "Rhode Island"
   
   sStates(37, 0) = "CT"
   sStates(37, 1) = "Connecticut"
   
   sStates(38, 0) = "NJ"
   sStates(38, 1) = "New Jersey"
   
   sStates(39, 0) = "DE"
   sStates(39, 1) = "Delaware"
   
   sStates(40, 0) = "NY"
   sStates(40, 1) = "New York"
   
   sStates(41, 0) = "PA"
   sStates(41, 1) = "Pennsylvania"
   
   sStates(42, 0) = "OH"
   sStates(42, 1) = "Ohio"
   
   sStates(43, 0) = "MD"
   sStates(43, 1) = "Maryland"
   
   sStates(44, 0) = "va"
   sStates(44, 1) = "Virginia"
   
   sStates(45, 0) = "nc"
   sStates(45, 1) = "North Carolina"
   
   sStates(46, 0) = "sc"
   sStates(46, 1) = "South Carolina"
   
   sStates(47, 0) = "ak"
   sStates(47, 1) = "Alaska"
   
   sStates(48, 0) = "hi"
   sStates(48, 1) = "Hawaii"
   
   sStates(49, 0) = "wv"
   sStates(49, 1) = "West Virginia"
   prg1.value = 20
   
   For iList = 0 To 49
      If prg1.value < 95 Then prg1.value = prg1.value + 2
      sSql = "INSERT INTO CsteTable (STATECODE,STATEDESC) " _
             & "VALUES('" & UCase$(sStates(iList, 0)) & "','" _
             & sStates(iList, 1) & "')"
      clsADOCon.ExecuteSQL sSql
   Next
   prg1.value = 100
   On Error Resume Next
   MouseCursor 0
   prg1.Visible = False
   lblSte.Visible = False
   FillCmbStates
   Exit Sub
   
DiaErr1:
   sProcName = "buildstate"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub FillCmbStates()
   On Error GoTo DiaErr1
   sSql = "SELECT STATECODE FROM CsteTable"
   LoadComboBox cmbSte, -1
   If cmbSte.ListCount > 0 Then
      cmbSte = cmbSte.List(0)
   Else
      BuildStateCodes
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "fillcmbst"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Function GetState() As Boolean
   sSql = "SELECT * FROM CsteTable WHERE STATECODE='" & cmbSte & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCde, ES_KEYSET)
   If bSqlRows Then
      With RdoCde
         txtDsc = "" & Trim(!STATEDESC)
         cmbReg = "" & Trim(!STATEREG)
         cmbDiv = "" & Trim(!STATEDIV)
         optDef.value = !STATEDEFAULT
      End With
      GetState = True
   Else
      txtDsc = ""
      optDef.value = vbUnchecked
      GetState = False
   End If
   
End Function

Private Sub optDef_LostFocus()
   On Error Resume Next
   If optDef.value = vbChecked Then
      sSql = "UPDATE CsteTable SET STATEDEFAULT=0 WHERE " _
             & "STATECODE<> '" & cmbSte & "' "
      clsADOCon.ExecuteSQL sSql
      
      sSql = "UPDATE CsteTable SET STATEDEFAULT=1 WHERE " _
             & "STATECODE='" & cmbSte & "' "
      clsADOCon.ExecuteSQL sSql
   Else
      sSql = "UPDATE CsteTable SET STATEDEFAULT=0 WHERE " _
             & "STATECODE='" & cmbSte & "' "
      clsADOCon.ExecuteSQL sSql
   End If
   
End Sub


Private Sub txtDsc_LostFocus()
   txtDsc = CheckLen(txtDsc, 30)
   txtDsc = StrCase(txtDsc)
   If bGoodState Then
      On Error Resume Next
      'RdoCde.Edit
      RdoCde!STATEDESC = "" & txtDsc
      RdoCde.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub



Private Sub AddCode()
   Dim bResponse As Byte
   Dim sMsg As String
   On Error GoTo DiaErr1
   sMsg = "That Code Wasn't Found." & vbCrLf _
          & "Are You Sure That You Want To Add It?"
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      sSql = "INSERT INTO CsteTable (STATECODE,STATEREG,STATEDIV) " _
             & "VALUES('" & cmbSte & "','" _
             & cmbReg & "','" & cmbDiv & "')"
      clsADOCon.ExecuteSQL sSql
      If clsADOCon.RowsAffected > 0 Then
         SysMsg "State Code Added.", True
         AddComboStr cmbSte.hwnd, cmbSte
         On Error Resume Next
         txtDsc.SetFocus
      End If
   Else
      CancelTrans
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "addcode"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub
