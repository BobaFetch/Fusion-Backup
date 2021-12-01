VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form ppiESf05a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change An Estimating Formula"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6285
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   6285
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdChg 
      Cancel          =   -1  'True
      Caption         =   "&Apply"
      Enabled         =   0   'False
      Height          =   315
      Left            =   5280
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "Change The Formula ID"
      Top             =   1920
      Width           =   875
   End
   Begin VB.TextBox txtFrm 
      Height          =   288
      Left            =   1920
      TabIndex        =   6
      Tag             =   "3"
      ToolTipText     =   "The New Formula ID"
      Top             =   1920
      Width           =   1572
   End
   Begin VB.ComboBox cmbFrm 
      Height          =   288
      Left            =   1920
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Formula Name (12) Characters Max"
      Top             =   720
      Width           =   1815
   End
   Begin VB.TextBox txtDsc 
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   1
      Tag             =   "2"
      Text            =   " "
      ToolTipText     =   "(30) Char Maximun"
      Top             =   1080
      Width           =   3912
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Close"
      Height          =   435
      Left            =   5280
      TabIndex        =   2
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
      FormDesignHeight=   2655
      FormDesignWidth =   6285
   End
   Begin VB.Label lblCenter 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1920
      TabIndex        =   9
      Top             =   1440
      Width           =   1812
   End
   Begin VB.Label Fr 
      BackStyle       =   0  'Transparent
      Caption         =   "Work Center"
      ForeColor       =   &H00400000&
      Height          =   252
      Index           =   3
      Left            =   360
      TabIndex        =   8
      Top             =   1440
      Width           =   1332
   End
   Begin VB.Label Fr 
      BackStyle       =   0  'Transparent
      Caption         =   "New Formula ID"
      ForeColor       =   &H00400000&
      Height          =   252
      Index           =   1
      Left            =   360
      TabIndex        =   5
      Top             =   1920
      Width           =   1332
   End
   Begin VB.Label Fr 
      BackStyle       =   0  'Transparent
      Caption         =   "Existing Formula ID"
      ForeColor       =   &H00400000&
      Height          =   252
      Index           =   2
      Left            =   360
      TabIndex        =   4
      Top             =   720
      Width           =   2292
   End
   Begin VB.Label Fr 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      ForeColor       =   &H00400000&
      Height          =   252
      Index           =   0
      Left            =   360
      TabIndex        =   3
      Top             =   1080
      Width           =   1332
   End
End
Attribute VB_Name = "ppiESf05a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
'12/7/05 New
'5/23/06 Added lblCenter and Enable/Disable cmdChg
Option Explicit
Dim bOnLoad As Byte
Dim bGoodFrm As Byte
Dim bBadNew As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Function GetFormula() As Byte
   Dim RdoForm As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT FORMULA_REF,FORMULA_DESC,FORMULA_CENTER " _
          & "FROM EsfrTable WHERE FORMULA_REF='" _
          & Compress(cmbFrm) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoForm, ES_FORWARD)
   If bSqlRows Then
      With RdoForm
         cmbFrm = "" & Trim(!FORMULA_REF)
         txtDsc = "" & Trim(!FORMULA_DESC)
         lblCenter = "" & Trim(!FORMULA_CENTER)
         GetFormula = 1
         .Cancel
      End With
      ClearResultSet RdoForm
   Else
      GetFormula = 0
   End If
   Set RdoForm = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getformula"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub cmbFrm_Click()
   bGoodFrm = GetFormula()
   
End Sub

Private Sub cmbFrm_LostFocus()
   cmbFrm = Compress(cmbFrm)
   cmbFrm = CheckLen(cmbFrm, 12)
   If Len(cmbFrm) Then
      bGoodFrm = GetFormula()
   Else
      txtDsc = ""
      bGoodFrm = 0
   End If
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub



Private Sub cmdChg_Click()
   Dim bGoodNew As Byte
   If Trim(txtFrm) = "" Then Exit Sub
   bGoodNew = TestFormula()
   If bGoodNew Then
      ChangeFormula
   Else
      MsgBox "That Formula ID Already Exists. Cannot Change.", _
         vbInformation, Caption
   End If
   
   
End Sub

Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then FillCombo
   bOnLoad = 0
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST
   FormatControls
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set ppiESf05a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtDsc.BackColor = BackColor
   
End Sub

Private Sub FillCombo()
   cmbFrm.Clear
   cmdChg.Enabled = False
   sSql = "SELECT FORMULA_REF FROM EsfrTable WHERE FORMULA_REF<>'NONE' " _
          & "ORDER BY FORMULA_REF"
   LoadComboBox cmbFrm, -1
   If cmbFrm.ListCount > 0 Then
      cmbFrm = cmbFrm.List(0)
      bGoodFrm = GetFormula()
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub




Private Function TestFormula() As Byte
   Dim RdoForm As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT FORMULA_REF FROM EsfrTable WHERE FORMULA_REF='" _
          & Compress(txtFrm) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoForm, ES_FORWARD)
   If bSqlRows Then
      TestFormula = 0
   Else
      TestFormula = 1
   End If
   ClearResultSet RdoForm
   Set RdoForm = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getformula"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub txtFrm_Change()
   If Len(Trim(txtFrm)) Then cmdChg.Enabled = True _
          Else cmdChg.Enabled = False
   
   
End Sub

Private Sub txtFrm_LostFocus()
   txtFrm = Compress(txtFrm)
   txtFrm = CheckLen(txtFrm, 12)
   
End Sub



Private Sub ChangeFormula()
   Dim bResponse As Byte
   Dim sMsg As String
   
   sMsg = "You Have Requested To Change Formula ID " & cmbFrm & "   " & vbCrLf _
          & "To " & txtFrm & " Would You Like to Continue?"
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      On Error Resume Next
      clsADOCon.BeginTrans
      sSql = "UPDATE EsrtTable SET BIDFORMULA='" & txtFrm & "' " _
             & "WHERE BIDFORMULA='" & cmbFrm & "'"
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
      
      sSql = "UPDATE EsfrTable SET FORMULA_REF='" & txtFrm & "' " _
             & "WHERE FORMULA_REF='" & cmbFrm & "'"
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
      If Err > 0 Then
         clsADOCon.RollbackTrans
         MsgBox "Could Not Make The Requested Change.", _
            vbInformation, Caption
      Else
         clsADOCon.CommitTrans
         SysMsg "The Formula Was Changed.", True
         txtFrm = " "
         FillCombo
      End If
   Else
      CancelTrans
   End If
End Sub
