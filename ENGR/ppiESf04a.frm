VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form ppiESf04a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add/Edit Formulae"
   ClientHeight    =   5805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6390
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   6390
   ShowInTaskbar   =   0   'False
   Begin VB.Frame z2 
      Height          =   30
      Left            =   240
      TabIndex        =   24
      Top             =   2040
      Width           =   6012
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "ppiESf04a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   23
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "Test"
      Enabled         =   0   'False
      Height          =   372
      Left            =   5280
      TabIndex        =   4
      Top             =   3120
      Width           =   972
   End
   Begin VB.TextBox txtFormula1 
      Height          =   288
      Left            =   1452
      TabIndex        =   5
      Tag             =   "1"
      Text            =   "Formula1"
      Top             =   3120
      Visible         =   0   'False
      Width           =   1032
   End
   Begin VB.TextBox txtFormula2 
      Height          =   288
      Left            =   1452
      TabIndex        =   6
      Tag             =   "1"
      Text            =   "Formula2"
      Top             =   3480
      Visible         =   0   'False
      Width           =   1032
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      Height          =   288
      Left            =   1452
      Locked          =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Tag             =   "1"
      Text            =   "Total"
      Top             =   4560
      Visible         =   0   'False
      Width           =   1032
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "Calculate"
      Height          =   372
      Left            =   2760
      TabIndex        =   9
      Top             =   4560
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.TextBox txtFormula3 
      Height          =   288
      Left            =   1452
      TabIndex        =   7
      Tag             =   "1"
      Text            =   "Formula3"
      Top             =   3840
      Visible         =   0   'False
      Width           =   1032
   End
   Begin VB.TextBox txtFormula4 
      Height          =   288
      Left            =   1452
      TabIndex        =   8
      Tag             =   "1"
      Text            =   "Formula4"
      Top             =   4200
      Visible         =   0   'False
      Width           =   1032
   End
   Begin VB.TextBox txtFormula 
      Enabled         =   0   'False
      Height          =   852
      Left            =   1440
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   2160
      Width           =   4812
   End
   Begin VB.TextBox txtDsc 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Tag             =   "2"
      Text            =   " "
      ToolTipText     =   "(30) Char Maximun"
      Top             =   960
      Width           =   3912
   End
   Begin VB.ComboBox cmbFrm 
      Height          =   288
      Left            =   1440
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Formula Name (12) Characters Max"
      Top             =   600
      Width           =   1815
   End
   Begin VB.ComboBox cmbWcn 
      Height          =   288
      Left            =   1440
      TabIndex        =   2
      Tag             =   "3"
      ToolTipText     =   "Work Center"
      Top             =   1320
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
      Left            =   7200
      Top             =   4080
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   5805
      FormDesignWidth =   6390
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   492
      Index           =   1
      Left            =   120
      TabIndex        =   22
      Top             =   5040
      Width           =   6132
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblFormula1 
      BackStyle       =   0  'Transparent
      Caption         =   "Formula1"
      Height          =   252
      Left            =   288
      TabIndex        =   21
      Top             =   3120
      Visible         =   0   'False
      Width           =   1212
   End
   Begin VB.Label lblFormula2 
      BackStyle       =   0  'Transparent
      Caption         =   "Formula2"
      Height          =   252
      Left            =   288
      TabIndex        =   20
      Top             =   3480
      Visible         =   0   'False
      Width           =   1212
   End
   Begin VB.Label lblTotal 
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      Height          =   252
      Left            =   288
      TabIndex        =   19
      Top             =   4560
      Visible         =   0   'False
      Width           =   1212
   End
   Begin VB.Label lblFormula3 
      BackStyle       =   0  'Transparent
      Caption         =   "Formula3"
      Height          =   252
      Left            =   288
      TabIndex        =   18
      Top             =   3840
      Visible         =   0   'False
      Width           =   1212
   End
   Begin VB.Label lblFormula4 
      BackStyle       =   0  'Transparent
      Caption         =   "Formula4"
      Height          =   252
      Left            =   288
      TabIndex        =   17
      Top             =   4200
      Visible         =   0   'False
      Width           =   1212
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Formula"
      Height          =   252
      Left            =   240
      TabIndex        =   15
      Top             =   2160
      Width           =   1092
   End
   Begin VB.Label lblWcn 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   1440
      TabIndex        =   14
      Top             =   1680
      Width           =   3132
   End
   Begin VB.Label Fr 
      BackStyle       =   0  'Transparent
      Caption         =   "Work Center"
      ForeColor       =   &H00400000&
      Height          =   252
      Index           =   1
      Left            =   240
      TabIndex        =   13
      Top             =   1320
      Width           =   1332
   End
   Begin VB.Label Fr 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      ForeColor       =   &H00400000&
      Height          =   252
      Index           =   0
      Left            =   240
      TabIndex        =   12
      Top             =   960
      Width           =   1332
   End
   Begin VB.Label Fr 
      BackStyle       =   0  'Transparent
      Caption         =   "Formula"
      ForeColor       =   &H00400000&
      Height          =   252
      Index           =   2
      Left            =   240
      TabIndex        =   11
      Top             =   600
      Width           =   1332
   End
End
Attribute VB_Name = "ppiESf04a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
Option Explicit
Dim ParseMath As New clsExpressionParser

Dim RdoForm As ADODB.Recordset
Dim bCancel As Byte
Dim bOnLoad As Byte
Dim bGoodFrm As Byte
Dim bGoodParse As Byte
Dim iVarCount As Integer
Dim sParseString(6) As String
Dim sCaptions(5) As String


Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Function VarCounter() As Integer
   Dim bVars As Byte
   Dim iRow As Integer
   Dim iBookMark As Integer
   Dim sExpress As String
   
   On Error Resume Next
   sExpress = Trim(txtFormula.Text)
   iBookMark = 1
   For iRow = 1 To 6
      iBookMark = InStr(iBookMark, sExpress, "VAR[")
      If iBookMark > 0 Then
         bVars = bVars + 1
         iBookMark = iBookMark + 1
      Else
         Exit For
      End If
   Next
   If bVars > 4 Then
      MsgBox "There Are More Than The (4) Variable Limit... ", _
         vbInformation, Caption
      VarCounter = 0
   Else
      VarCounter = bVars
   End If
   
End Function

Private Function ParseFormulaString() As Byte
   Dim iList As Integer
   Dim iBookMark1 As Integer
   Dim iBookMark2 As Integer
   Dim iBookMark3 As Integer
   Dim sMathString As String
   
   cmdCalc.Visible = False
   txtFormula1.Visible = False
   txtFormula2.Visible = False
   txtFormula3.Visible = False
   txtFormula4.Visible = False
   lblFormula1.Visible = False
   lblFormula2.Visible = False
   lblFormula3.Visible = False
   lblFormula4.Visible = False
   
   txtTotal.Visible = False
   lblTotal.Visible = False
   txtFormula1.Text = ""
   txtFormula2.Text = ""
   txtFormula3.Text = ""
   txtFormula4.Text = ""
   Label3(1) = ""
   sMathString = Trim(txtFormula.Text)
   
   Erase sParseString
   Erase sCaptions
   
   'Get the expression
   On Error GoTo DiaErr1
   If iVarCount > 0 Then
      iBookMark1 = 1
      For iList = 1 To iVarCount
         iBookMark2 = InStr(iBookMark1, sMathString, "VAR[")
         If iBookMark2 > 1 Then
            sParseString(iList) = Mid$(sMathString, iBookMark1, (iBookMark2 - iBookMark1))
            iBookMark1 = InStr(iBookMark2, sMathString, "]") + 1
         Else
            If iBookMark2 = 1 Then
               iBookMark1 = InStr(iBookMark2, sMathString, "]") + 1
            End If
         End If
      Next
      sParseString(iList) = Right$(sMathString, Len(sMathString) - (iBookMark1 - 1))
      'get the captions
      iBookMark1 = 1
      For iList = 1 To iVarCount
         iBookMark2 = InStr(iBookMark1, sMathString, "VAR[")
         If iBookMark2 > 0 Then
            iBookMark1 = InStr(iBookMark2, sMathString, "]")
            sCaptions(iList) = Mid$(sMathString, iBookMark2 + 4, iBookMark1 - (iBookMark2 + 4))
         End If
      Next
   End If
   
   txtTotal = "0.00"
   Select Case iVarCount
      Case 1
         lblFormula1.Visible = True
         lblFormula1.Caption = sCaptions(1)
         txtFormula1.Visible = True
         txtFormula1.Text = "0.000"
         lblTotal.Top = lblFormula1.Top + lblFormula1.Height + 200
         txtTotal.Top = lblTotal.Top
         lblTotal.Visible = True
         txtTotal.Visible = True
         cmdCalc.Top = txtTotal.Top
         cmdCalc.Visible = True
      Case 2
         lblFormula1.Visible = True
         lblFormula1.Caption = sCaptions(1)
         txtFormula1.Visible = True
         txtFormula1.Text = "0.000"
         
         lblFormula2.Visible = True
         lblFormula2.Caption = sCaptions(2)
         txtFormula2.Visible = True
         txtFormula2.Text = "0.000"
         
         lblTotal.Top = lblFormula2.Top + lblFormula2.Height + 200
         txtTotal.Top = lblTotal.Top
         lblTotal.Visible = True
         txtTotal.Visible = True
         cmdCalc.Top = txtTotal.Top
         cmdCalc.Visible = True
      Case 3
         lblFormula1.Visible = True
         lblFormula1.Caption = sCaptions(1)
         txtFormula1.Visible = True
         txtFormula1.Text = "0.000"
         
         lblFormula2.Visible = True
         lblFormula2.Caption = sCaptions(2)
         txtFormula2.Visible = True
         txtFormula2.Text = "0.000"
         
         lblFormula3.Visible = True
         lblFormula3.Caption = sCaptions(3)
         txtFormula3.Visible = True
         txtFormula3.Text = "0.000"
         
         
         lblTotal.Top = lblFormula3.Top + lblFormula3.Height + 200
         txtTotal.Top = lblTotal.Top
         lblTotal.Visible = True
         txtTotal.Visible = True
         cmdCalc.Top = txtTotal.Top
         cmdCalc.Visible = True
      Case 4
         lblFormula1.Visible = True
         lblFormula1.Caption = sCaptions(1)
         txtFormula1.Visible = True
         txtFormula1.Text = "0.000"
         
         lblFormula2.Visible = True
         lblFormula2.Caption = sCaptions(2)
         txtFormula2.Visible = True
         txtFormula2.Text = "0.000"
         
         lblFormula3.Visible = True
         lblFormula3.Caption = sCaptions(3)
         txtFormula3.Visible = True
         txtFormula3.Text = "0.000"
         
         lblFormula4.Visible = True
         lblFormula4.Caption = sCaptions(4)
         txtFormula4.Visible = True
         txtFormula4.Text = "0.000"
         
         lblTotal.Top = lblFormula4.Top + lblFormula4.Height + 200
         txtTotal.Top = lblTotal.Top
         lblTotal.Visible = True
         txtTotal.Visible = True
         cmdCalc.Top = txtTotal.Top
         cmdCalc.Visible = True
   End Select
   ParseFormulaString = 1
   Exit Function
   
DiaErr1:
   ParseFormulaString = 0
   
End Function


Private Sub cmbFrm_Click()
   bGoodFrm = GetFormula()
   
End Sub


Private Sub cmbFrm_LostFocus()
   cmbFrm = Compress(cmbFrm)
   cmbFrm = CheckLen(cmbFrm, 12)
   If bCancel = 0 Then
      If Len(cmbFrm) Then
         bGoodFrm = GetFormula()
         If bGoodFrm = 0 Then AddFormula
      Else
         bGoodFrm = 0
      End If
   End If
   
End Sub


Private Sub cmbWcn_Click()
   GetThisWorkCenter
   
End Sub


Private Sub cmbWcn_LostFocus()
   cmbWcn = CheckLen(cmbWcn, 10)
   GetThisWorkCenter
   If bGoodFrm Then
      On Error Resume Next
      With RdoForm
         '.Edit
         If lblWcn.ForeColor = ES_RED Then
            !FORMULA_CENTER = ""
         Else
            !FORMULA_CENTER = Compress(cmbWcn)
         End If
         .Update
      End With
   End If
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub



Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   bCancel = 1
   
End Sub



Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 8513
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub

Private Sub cmdTest_Click()
   txtFormula.Text = UCase$(txtFormula.Text)
   txtFormula.Text = Replace(txtFormula, "'", "")
   txtFormula.Text = Replace(txtFormula, Chr$(34), "")
   txtFormula.Text = Replace(txtFormula, "x", "*")
   txtFormula.Text = Replace(txtFormula, "X", "*")
   txtFormula.Text = Replace(txtFormula, " ", "")
   iVarCount = VarCounter()
   ES_SYSDATE = GetServerDateTime()
   bGoodParse = ParseFormulaString()
   If bGoodFrm Then
      On Error Resume Next
      With RdoForm
         '.Edit
         !FORMULA_REVISED = Format(ES_SYSDATE, "mm/dd/yy hh:mm")
         !FORMULA_REVISEDBY = sInitials
         !FORMULA_TEXT = txtFormula.Text
         !FORMULA_VARIABLES = iVarCount
         .Update
      End With
   End If
   If bGoodParse = 0 Then MsgBox "The Formula Would Not Parse.", _
                   vbInformation, Caption
   
End Sub


Private Sub cmdCalc_Click()
   Dim bByte As Byte
   Dim dValue As Currency
   Dim sString As String
   
   bByte = 0
   If txtFormula1.Visible And Val(txtFormula1) = 0 Then bByte = 1
   If txtFormula2.Visible And Val(txtFormula2) = 0 Then bByte = 1
   If txtFormula3.Visible And Val(txtFormula3) = 0 Then bByte = 1
   If txtFormula4.Visible And Val(txtFormula4) = 0 Then bByte = 1
   
   If bByte = 0 Then
      sString = sParseString(1) & txtFormula1 & sParseString(2) _
                & txtFormula2 & sParseString(3) _
                & txtFormula3 & sParseString(4) _
                & txtFormula4 & sParseString(5)
      dValue = ParseMath.ParseExpression(sString)
      txtTotal = Format(dValue, "####0.00")
      Label3(1) = "The Actual Formula Is: " & sString & "= " & txtTotal.Text
   Else
      MsgBox "Each Entry Requires A Number More Than Zero.", _
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
   FormLoad Me
   FormatControls
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set RdoForm = Nothing
   Set ppiESf04a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub FillCombo()
   On Error GoTo DiaErr1
   sSql = "SELECT DISTINCT WCNREF,WCNNUM FROM WcntTable ORDER BY WCNREF"
   LoadComboBox cmbWcn
   
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

Private Function GetFormula() As Byte
   CloseBoxes
   On Error GoTo DiaErr1
   sSql = "SELECT FORMULA_REF,FORMULA_DESC,FORMULA_CENTER," _
          & "FORMULA_REVISED,FORMULA_REVISEDBY,FORMULA_VARIABLES," _
          & "FORMULA_TEXT FROM EsfrTable WHERE FORMULA_REF='" _
          & Compress(cmbFrm) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoForm, ES_KEYSET)
   If bSqlRows Then
      With RdoForm
         cmbFrm = "" & Trim(!FORMULA_REF)
         txtDsc = "" & Trim(!FORMULA_DESC)
         cmbWcn = "" & Trim(!FORMULA_CENTER)
         txtFormula = "" & Trim(!FORMULA_TEXT)
         GetThisWorkCenter
         GetFormula = 1
      End With
      cmdTest.Enabled = True
      txtFormula.Enabled = True
   Else
      GetFormula = 0
   End If
   Exit Function
   
DiaErr1:
   sProcName = "getformula"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub AddFormula()
   Dim bResponse As Byte
   
   On Error Resume Next
   bResponse = MsgBox("Formula " & cmbFrm & " Wasn't Found. Add A New Formula?", _
               ES_YESQUESTION, Caption)
   If bResponse = vbNo Then
      CancelTrans
   Else
      bResponse = IllegalCharacters(cmbFrm)
      If bResponse > 0 Then
         MsgBox "The Part Number Contains An Illegal " & Chr$(bResponse) & ".", _
            vbExclamation, Caption
      Else
         'Add one
         sSql = "INSERT INTO EsfrTable (FORMULA_REF,FORMULA_CREATEDBY," _
                & "FORMULA_REVISEDBY) VALUES('" & cmbFrm & "','" _
                & sInitials & "','" & sInitials & "')"
         clsADOCon.ExecuteSQL sSql 'rdExecDirect
         If Err = 0 Then
            SysMsg "Formula Was added.", True
            bGoodFrm = GetFormula()
         Else
            MsgBox "Couldn't Add Formula.", vbExclamation, Caption
         End If
      End If
   End If
End Sub

Private Sub lblWcn_Change()
   If Left(lblWcn, 8) = "*** Requ" Then _
           lblWcn.ForeColor = ES_RED Else lblWcn.ForeColor = vbBlack
   
End Sub

Private Sub txtDsc_LostFocus()
   txtDsc = CheckLen(txtDsc, 40)
   txtDsc = StrCase(txtDsc)
   If bGoodFrm Then
      On Error Resume Next
      With RdoForm
         '.Edit
         !FORMULA_DESC = Trim(txtDsc)
         .Update
      End With
   End If
End Sub



Private Sub GetThisWorkCenter()
   Dim RdoWrk As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT WCNREF,WCNNUM,WCNDESC FROM WcntTable WHERE WCNREF='" _
          & Compress(cmbWcn) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoWrk, ES_FORWARD)
   If bSqlRows Then
      cmbWcn = "" & Trim(RdoWrk!WCNNUM)
      lblWcn = "" & Trim(RdoWrk!WCNDESC)
   Else
      lblWcn = "*** Requires A Valid Work Center ***"
   End If
   Set RdoWrk = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getthiswork"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub txtFormula1_LostFocus()
   txtFormula1 = Format(Abs(Val(txtFormula1)), "#####0.000")
   
End Sub


Private Sub txtFormula2_LostFocus()
   txtFormula2 = Format(Abs(Val(txtFormula2)), "#####0.000")
   
End Sub


Private Sub txtFormula3_LostFocus()
   txtFormula3 = Format(Abs(Val(txtFormula3)), "#####0.000")
   
End Sub


Private Sub txtFormula4_LostFocus()
   txtFormula4 = Format(Abs(Val(txtFormula4)), "#####0.000")
   
End Sub



Public Sub CloseBoxes()
   cmdTest.Enabled = False
   cmdCalc.Visible = False
   txtFormula.Enabled = False
   txtFormula1.Visible = False
   txtFormula2.Visible = False
   lblFormula1.Visible = False
   lblFormula2.Visible = False
   txtTotal.Visible = False
   lblTotal.Visible = False
   txtFormula1.Text = ""
   txtFormula2.Text = ""
   txtFormula3.Text = ""
   txtFormula4.Text = ""
   
End Sub
