VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form SadmSLe06a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Country Codes"
   ClientHeight    =   2685
   ClientLeft      =   1200
   ClientTop       =   855
   ClientWidth     =   5130
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   5130
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "SadmSLe06a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   4800
      Top             =   2640
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2685
      FormDesignWidth =   5130
   End
   Begin VB.TextBox txtDsc 
      Height          =   285
      Left            =   1360
      TabIndex        =   1
      Top             =   960
      Width           =   3495
   End
   Begin VB.TextBox txtCmt 
      Height          =   975
      Left            =   1360
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1320
      Width           =   3475
   End
   Begin VB.ComboBox cmbTrm 
      Height          =   315
      Left            =   1360
      Sorted          =   -1  'True
      TabIndex        =   0
      ToolTipText     =   "Enter/Revise Terms (2 char)"
      Top             =   600
      Width           =   2000
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   4200
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Comment"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Country Code"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "SadmSLe06a"
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

Dim AdoQry As ADODB.Command
Dim AdoParameter As ADODB.Parameter
Dim RdoCnt As ADODB.Recordset

Dim bOnLoad As Byte
Dim bGoodCountry As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd



Private Sub cmbTrm_Click()
   bGoodCountry = GetCountry()
   
End Sub

Private Sub cmbTrm_LostFocus()
   cmbTrm = CheckLen(cmbTrm, 15)
   If Len(cmbTrm) Then
      cmbTrm = Compress(cmbTrm)
      bGoodCountry = GetCountry()
      If bGoodCountry = 0 Then AddCountry
   Else
      bGoodCountry = 0
   End If
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   cmbTrm = ""
   
End Sub


Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 1206
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub



Private Sub Form_Activate()
   If bOnLoad Then
      FillCombo
      bOnLoad = 0
      If cmbTrm.ListCount > 0 Then cmbTrm = cmbTrm.List(0)
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   
   sSql = "SELECT TOP 1 * FROM " _
          & "CntrTable WHERE COUREF= ? "
          
   Set AdoQry = New ADODB.Command
   AdoQry.CommandText = sSql
   
   Set AdoParameter = New ADODB.Parameter
   AdoParameter.Size = 15
   AdoParameter.Type = adChar
   
   AdoQry.Parameters.Append AdoParameter
          
   'Set RdoQry = RdoCon.CreateQuery("", sSql)
   'RdoQry.MaxRows = 1
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set SadmSLe06a = Nothing
   
End Sub




Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
   
End Sub



Private Sub txtCmt_LostFocus()
   txtCmt = CheckLen(txtCmt, 255)
   txtCmt = StrCase(txtCmt, ES_FIRSTWORD)
   On Error Resume Next
   If bGoodCountry Then
      'RdoCnt.Edit
      RdoCnt!COUCOMT = "" & txtCmt
      RdoCnt.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtDsc_LostFocus()
   txtDsc = CheckLen(txtDsc, 40)
   txtDsc = StrCase(txtDsc)
   On Error Resume Next
   If bGoodCountry Then
      'RdoCnt.Edit
      RdoCnt!COUDESC = "" & txtDsc
      RdoCnt.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub



Private Function GetCountry() As Byte
   Dim sTerms As String
   sTerms = cmbTrm
   On Error GoTo DiaErr1
   'RdoQry(0) = sTerms
   AdoQry.Parameters(0).value = sTerms
   bSqlRows = clsADOCon.GetQuerySet(RdoCnt, AdoQry, ES_KEYSET, True)
   If bSqlRows Then
      With RdoCnt
         cmbTrm = "" & Trim(!COUCOUNTRY)
         txtDsc = "" & Trim(!COUDESC)
         txtCmt = "" & Trim(!COUCOMT)
      End With
      GetCountry = 1
   Else
      txtDsc = ""
      txtCmt = ""
      GetCountry = 0
   End If
   Exit Function
   
DiaErr1:
   sProcName = "getterms"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub AddCountry()
   Dim bResponse As Byte
   Dim sMsg As String
   Dim sTerms As String
   
   sTerms = Compress(cmbTrm)
   sMsg = sTerms & " Wasn't Found. Add The Country Code?"
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      On Error GoTo DiaErr1
      sSql = "INSERT INTO CntrTable (COUREF,COUCOUNTRY) " _
             & "VALUES('" & sTerms & "','" & cmbTrm & "')"
      clsADOCon.ExecuteSQL sSql
      If clsADOCon.RowsAffected Then
         SysMsg "Country Added.", True
         cmbTrm = sTerms
         AddComboStr cmbTrm.hwnd, sTerms
         bGoodCountry = GetCountry()
         On Error Resume Next
         txtDsc.SetFocus
      Else
         MsgBox "Couldn't Add The Terms.", vbExclamation, Caption
      End If
   Else
      CancelTrans
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "addterms"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub FillCombo()
   On Error GoTo DiaErr1
   sSql = "Qry_FillCountries"
   LoadComboBox cmbTrm
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub
