VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form AdmnQAe02a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Vendor Information"
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7335
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   7335
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "AdmnQAe02a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   28
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox cmbSte 
      Height          =   315
      Left            =   4560
      TabIndex        =   5
      Tag             =   "3"
      Top             =   2160
      Width           =   735
   End
   Begin VB.TextBox txtZip 
      Height          =   285
      Left            =   5880
      TabIndex        =   6
      Tag             =   "1"
      Top             =   2160
      Width           =   1335
   End
   Begin VB.TextBox txtCty 
      Height          =   285
      Left            =   1440
      TabIndex        =   4
      Tag             =   "2"
      Top             =   2160
      Width           =   2085
   End
   Begin VB.TextBox txtAdr 
      Height          =   855
      Left            =   1440
      MultiLine       =   -1  'True
      TabIndex        =   3
      Tag             =   "9"
      Top             =   1200
      Width           =   3475
   End
   Begin VB.TextBox txtEml 
      Height          =   285
      Left            =   1440
      TabIndex        =   11
      Tag             =   "2"
      ToolTipText     =   "Click Here To Send E-Mail (Requires An Entry)"
      Top             =   3240
      Width           =   5340
   End
   Begin VB.TextBox txtExt 
      Height          =   285
      Left            =   3600
      TabIndex        =   8
      Tag             =   "1"
      Top             =   2520
      Width           =   735
   End
   Begin VB.TextBox txtCnt 
      Height          =   285
      Left            =   1440
      TabIndex        =   10
      Tag             =   "2"
      Top             =   2880
      Width           =   2535
   End
   Begin VB.TextBox txtByr 
      Height          =   285
      Left            =   1440
      TabIndex        =   12
      Tag             =   "2"
      Top             =   3600
      Width           =   2535
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   6360
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   0
      Width           =   915
   End
   Begin VB.TextBox txtNme 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Tag             =   "2"
      Top             =   840
      Width           =   3475
   End
   Begin VB.ComboBox cmbVnd 
      Height          =   315
      Left            =   1440
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Select Vendor From List"
      Top             =   480
      Width           =   1555
   End
   Begin VB.TextBox txtTyp 
      Height          =   285
      Left            =   6240
      TabIndex        =   2
      Tag             =   "3"
      Top             =   840
      Width           =   495
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6360
      Top             =   3600
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4110
      FormDesignWidth =   7335
   End
   Begin MSMask.MaskEdBox txtPhn 
      Height          =   285
      Left            =   1440
      TabIndex        =   7
      Top             =   2520
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   12
      Mask            =   "###-###-####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtFax 
      Height          =   285
      Left            =   5040
      TabIndex        =   9
      Top             =   2520
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   12
      Mask            =   "###-###-####"
      PromptChar      =   "_"
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Fax"
      Height          =   285
      Index           =   7
      Left            =   4560
      TabIndex        =   27
      Top             =   2520
      Width           =   705
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Zip"
      Height          =   285
      Index           =   22
      Left            =   5400
      TabIndex        =   26
      Top             =   2160
      Width           =   555
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "State"
      Height          =   285
      Index           =   4
      Left            =   3720
      TabIndex        =   25
      Top             =   2160
      Width           =   675
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "City"
      Height          =   285
      Index           =   3
      Left            =   120
      TabIndex        =   24
      Top             =   2160
      Width           =   915
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Phone"
      Height          =   285
      Index           =   6
      Left            =   120
      TabIndex        =   23
      Top             =   2520
      Width           =   915
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   22
      Top             =   1200
      Width           =   1035
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "E-Mail"
      Height          =   285
      Index           =   5
      Left            =   120
      TabIndex        =   21
      Top             =   3240
      Width           =   915
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ext"
      Height          =   285
      Index           =   8
      Left            =   3120
      TabIndex        =   20
      Top             =   2520
      Width           =   435
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Contact"
      Height          =   285
      Index           =   9
      Left            =   120
      TabIndex        =   19
      Top             =   2880
      Width           =   1035
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Buyer"
      Height          =   285
      Index           =   33
      Left            =   120
      TabIndex        =   18
      Top             =   3600
      Width           =   915
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Vendor Name"
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   17
      Top             =   840
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nickname"
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   16
      Top             =   480
      Width           =   1425
   End
   Begin VB.Label lblNum 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3120
      TabIndex        =   15
      Top             =   480
      Width           =   375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Vendor Type"
      Height          =   285
      Index           =   32
      Left            =   5100
      TabIndex        =   14
      Top             =   840
      Width           =   1065
   End
End
Attribute VB_Name = "AdmnQAe02a"
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
Dim AdoVdr As ADODB.Recordset

Dim bCanceled As Boolean
Dim bGoodVendor As Byte
Dim bOnLoad As Byte

Dim sVendRef As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbSte_LostFocus()
   cmbSte = CheckLen(cmbSte, 2)
   If bGoodVendor Then
      On Error Resume Next
      AdoVdr!VEBSTATE = cmbSte
      AdoVdr.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub cmbVnd_Click()
   bGoodVendor = GetVendor()
   
End Sub


Private Sub cmbVnd_LostFocus()
   cmbVnd = CheckLen(cmbVnd, 10)
   If bCanceled Then Exit Sub
   If Len(cmbVnd) Then
      bGoodVendor = GetVendor()
      If bGoodVendor = 0 Then AddVendor
   End If
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub


Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   bCanceled = True
   
End Sub


Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 6502
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillVendors Me
      If cUR.CurrentVendor <> "" Then cmbVnd = cUR.CurrentVendor
      If cmbVnd.ListCount > 0 Then bGoodVendor = GetVendor()
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Function GetVendor() As Byte
   ClearBoxes
   On Error GoTo DiaErr1
   'RdoQry(0) = Compress(cmbVnd)
   AdoQry.Parameters(0).Value = Compress(cmbVnd)
   bSqlRows = clsADOCon.GetQuerySet(AdoVdr, AdoQry, ES_KEYSET, True)
   If bSqlRows Then
      With AdoVdr
         GetVendor = 1
         txtPhn.Mask = ""
         txtFax.Mask = ""
         cmbVnd = "" & Trim(!VENICKNAME)
         lblNum = 0 + Format(!VENUMBER, "#0")
         txtNme = "" & Trim(!VEBNAME)
         txtTyp = "" & Trim(!VETYPE)
         txtAdr = "" & Trim(!VEBADR)
         txtCty = "" & Trim(!VEBCITY)
         cmbSte = "" & Trim(!VEBSTATE)
         txtZip = "" & Trim(!VEBZIP)
         txtPhn = "" & Trim(!VEBPHONE)
         txtCnt = "" & Trim(!VEBCONTACT)
         txtFax = "" & Trim(!VEFAX)
         If !VEBEXT > 0 Then txtExt = !VEBEXT
         txtEml = "" & Trim(!VEEMAIL)
         txtByr = "" & Trim(!VEBUYER)
         txtPhn.Mask = "###-###-####"
         txtFax.Mask = "###-###-####"
         Err.Clear
      End With
   Else
      txtNme = "*** No Current Vendor ***"
      GetVendor = 0
   End If
   Exit Function
   
DiaErr1:
   sProcName = "getvendor"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub ClearBoxes()
   Dim iList As Integer
   For iList = 0 To Controls.Count - 1
      If TypeOf Controls(iList) Is TextBox Then
         Controls(iList).Text = " "
      End If
   Next
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   
   sSql = "SELECT TOP 1 VEREF,VENICKNAME,VENUMBER,VEBNAME,VETYPE," _
          & "VEBADR,VEBCITY,VEBSTATE,VEBZIP,VEBPHONE,VEBCONTACT," _
          & "VEFAX,VEBEXT,VEEMAIL,VEBUYER FROM VndrTable " _
          & "WHERE VEREF= ? "
'   Set RdoQry = RdoCon.CreateQuery("", sSql)
'   RdoQry.MaxRows = 1
       
   Set AdoQry = New ADODB.Command
   AdoQry.CommandText = sSql
   
   Set AdoParameter = New ADODB.Parameter
   AdoParameter.Size = 10
   AdoParameter.Type = adChar
   
   AdoQry.Parameters.Append AdoParameter
   
   bOnLoad = 1
   
End Sub


Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   On Error Resume Next
   Set AdoParameter = Nothing
   Set AdoQry = Nothing
   Set AdoVdr = Nothing
   Set AdmnQAe02a = Nothing
   
End Sub


Private Sub txtAdr_LostFocus()
   txtAdr = CheckLen(txtAdr, 160)
   If bGoodVendor Then
      On Error Resume Next
      AdoVdr!VEBADR = txtAdr
      AdoVdr.Update
      If Err > 0 Then ValidateEdit
   End If
   
   
End Sub


Private Sub txtByr_LostFocus()
   txtByr = CheckLen(txtByr, 20)
   txtByr = StrCase(txtByr)
   If bGoodVendor Then
      On Error Resume Next
      AdoVdr!VEBUYER = "" & txtByr
      AdoVdr.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtCnt_LostFocus()
   txtCnt = CheckLen(txtCnt, 20)
   txtCnt = StrCase(txtCnt)
   If bGoodVendor Then
      On Error Resume Next
      AdoVdr!VEBCONTACT = txtCnt
      AdoVdr.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtCty_LostFocus()
   txtCty = CheckLen(txtCty, 18)
   txtCty = StrCase(txtCty)
   If bGoodVendor Then
      On Error Resume Next
      AdoVdr!VEBCITY = txtCty
      AdoVdr.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtEml_Click()
   If Trim(txtEml) <> "" Then SendEMail Trim(txtEml)
   
End Sub

Private Sub txtEml_LostFocus()
   txtEml = CheckLen(txtEml, 60)
   If bGoodVendor Then
      On Error Resume Next
      AdoVdr!VEEMAIL = txtEml
      AdoVdr.Update
      If Err > 0 Then ValidateEdit
   End If
   
   
End Sub


Private Sub txtExt_LostFocus()
   txtExt = CheckLen(txtExt, 4)
   If Len(txtExt) > 0 Then txtExt = Format(Abs(Val(txtExt)), "###0")
   If bGoodVendor Then
      On Error Resume Next
      AdoVdr!VEBEXT = Val(txtExt)
      AdoVdr.Update
      If Err > 0 Then ValidateEdit
   End If
   
   
End Sub


Private Sub txtFax_LostFocus()
   If bGoodVendor Then
      On Error Resume Next
      AdoVdr!VEFAX = "" & txtFax
      AdoVdr.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtNme_Change()
   If Left(txtNme, 9) = "*** No Cu" Then
      txtNme.ForeColor = ES_RED
      txtNme.BackColor = vbButtonFace
   Else
      txtNme.ForeColor = vbBlack
      txtNme.BackColor = vbWhite
   End If
   
End Sub

Private Sub txtNme_LostFocus()
   txtNme = CheckLen(txtNme, 40)
   txtNme = StrCase(txtNme)
   If bGoodVendor Then
      On Error Resume Next
      AdoVdr!VEBNAME = txtNme
      AdoVdr.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtPhn_LostFocus()
   If bGoodVendor Then
      On Error Resume Next
      AdoVdr!VEBPHONE = "" & txtPhn
      AdoVdr.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub




Private Sub txtTyp_LostFocus()
   txtTyp = CheckLen(txtTyp, 2)
   If bGoodVendor Then
      On Error Resume Next
      AdoVdr!VETYPE = "" & txtTyp
      AdoVdr.Update
      If Err > 0 Then ValidateEdit
   End If
   
   
End Sub


Private Sub txtZip_LostFocus()
   txtZip = CheckLen(txtZip, 10)
   If bGoodVendor Then
      On Error Resume Next
      AdoVdr!VEBZIP = "" & txtZip
      AdoVdr.Update
      If Err > 0 Then ValidateEdit
   End If
   
   
End Sub



Private Sub AddVendor()
   Dim bResponse As Byte
   Dim lNextVend As Long
   Dim sMsg As String
   Dim sVend As String
   
   bResponse = IllegalCharacters(cmbVnd)
   If bResponse > 0 Then
      MsgBox "The Vendor Nickname Contains An Illegal " & Chr$(bResponse) & ".", _
         vbExclamation, Caption
      Exit Sub
   End If
   lNextVend = GetNextNumber()
   On Error GoTo DiaErr1
   sMsg = cmbVnd & " Hasn't Been Established." & vbCr _
          & "Add The New Vendor?...        "
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      sVend = Compress(cmbVnd)
      If sVend = "ALL" Then
         MsgBox "ALL Is An Illegal Vendor Nickname.", _
            vbExclamation, Caption
         Exit Sub
      Else
         On Error Resume Next
         clsADOCon.ADOErrNum = 0
         sSql = "INSERT INTO VndrTable (VEREF,VENICKNAME,VENUMBER,VEBSTATE) " _
                & "VALUES ('" & sVend & "','" & Trim(cmbVnd) & "'," _
                & lNextVend & ",'" & cmbSte & "')"
         clsADOCon.BeginTrans
         clsADOCon.ExecuteSQL sSql
         If clsADOCon.ADOErrNum = 0 Then
            clsADOCon.CommitTrans
            MsgBox "The Vendor Was Successfully Added.", _
               vbInformation, Caption
            AddComboStr cmbVnd.hwnd, cmbVnd
            bGoodVendor = GetVendor()
         Else
            bGoodVendor = False
            clsADOCon.RollbackTrans
            clsADOCon.ADOErrNum = 0
            MsgBox "Couldn't Successfully Add The Vendor.", _
               vbInformation, Caption
         End If
      End If
   Else
      CancelTrans
      On Error Resume Next
      cmbVnd.SetFocus
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "addvendor"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Function GetNextNumber() As Long
   Dim RdoNxt As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT MAX(VENUMBER) FROM VndrTable"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoNxt, ES_FORWARD)
   If bSqlRows Then
      If Not IsNull(RdoNxt.Fields(0)) Then
         GetNextNumber = RdoNxt.Fields(0) + 1
      Else
         GetNextNumber = 1
      End If
      ClearResultSet RdoNxt
   Else
      GetNextNumber = 1
   End If
   Set RdoNxt = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getnextnum"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function
