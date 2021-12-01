VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Begin VB.Form AdmnUsrpw
   BorderStyle = 3 'Fixed Dialog
   Caption = "Change A User Password"
   ClientHeight = 2340
   ClientLeft = 48
   ClientTop = 336
   ClientWidth = 5892
   ClipControls = 0 'False
   ControlBox = 0 'False
   ForeColor = &H8000000F&
   LinkTopic = "Form1"
   LockControls = -1 'True
   MaxButton = 0 'False
   MDIChild = -1 'True
   MinButton = 0 'False
   Moveable = 0 'False
   ScaleHeight = 2340
   ScaleWidth = 5892
   ShowInTaskbar = 0 'False
   Begin VB.TextBox txtVpw
      Height = 285
      IMEMode = 3 'DISABLE
      Left = 1920
      PasswordChar = "*"
      TabIndex = 1
      ToolTipText = "Verify Password"
      Top = 1680
      Width = 1600
   End
   Begin VB.TextBox txtPas
      Height = 285
      IMEMode = 3 'DISABLE
      Left = 1920
      PasswordChar = "*"
      TabIndex = 0
      ToolTipText = "Case Sensitive Max (15) Char"
      Top = 1320
      Width = 1600
   End
   Begin VB.CommandButton cmdChg
      Caption = "&Apply"
      Enabled = 0 'False
      Height = 315
      Left = 4920
      TabIndex = 2
      ToolTipText = "Update To The Current User Password"
      Top = 600
      Width = 875
   End
   Begin VB.CommandButton cmdCan
      Cancel = -1 'True
      Caption = "Close"
      Height = 435
      Left = 4920
      TabIndex = 3
      TabStop = 0 'False
      Top = 0
      Width = 875
   End
   Begin ResizeLibCtl.ReSize ReSize1
      Left = 5880
      Top = 2040
      _Version = 196615
      _ExtentX = 593
      _ExtentY = 593
      _StockProps = 0
      Enabled = -1 'True
      FormMinWidth = 0
      FormMinHeight = 0
      FormDesignHeight = 2340
      FormDesignWidth = 5892
   End
   Begin VB.Label lblName
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      Height = 285
      Left = 1920
      TabIndex = 9
      Top = 960
      Width = 2655
   End
   Begin VB.Label lblUser
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      Height = 285
      Left = 1920
      TabIndex = 8
      Top = 600
      Width = 2055
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Verify Password"
      Height = 255
      Index = 3
      Left = 360
      TabIndex = 7
      Top = 1680
      Width = 1455
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Password"
      Height = 255
      Index = 2
      Left = 360
      TabIndex = 6
      Top = 1320
      Width = 1455
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "User Name"
      Height = 255
      Index = 1
      Left = 360
      TabIndex = 5
      Top = 960
      Width = 1335
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "User Id"
      Height = 255
      Index = 0
      Left = 360
      TabIndex = 4
      Top = 600
      Width = 1335
   End
End
Attribute VB_Name = "AdmnUsrpw"
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

Dim bOnLoad As Byte

Dim sNewPass As String
Dim sVerify As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmdCan_Click()
   Unload Me
   
End Sub



Private Sub cmdChg_Click()
   sNewPass = Trim(txtPas)
   sVerify = Trim(txtVpw)
   On Error Resume Next
   If Len(sNewPass) < 3 Then
      MsgBox "The Password Must Be At Least (3) Characters.", _
         vbInformation, Caption
      txtPas.SetFocus
      Exit Sub
   End If
   
   If sNewPass <> sVerify Then
      MsgBox "The Password Must Be The Same As The Verification.", _
         vbInformation, Caption
      txtPas.SetFocus
      Exit Sub
   End If
   cmdCan.Enabled = False
   sNewPass = ScramblePw(sNewPass)
   SecPw.PassWord = sNewPass
   Put #iFreeIdx, iCurrentRec, SecPw
   If Err = 0 Then
      SysMsg "User Password Updated.", True
   Else
      MsgBox "Could Not Change The Password.", _
         vbExclamation, Caption
   End If
   Sleep 1000
   Unload Me
   
End Sub

Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      FillUser
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST
   Move 800, 800
   FormatControls
   
   bOnLoad = 1
   Show
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   AdmnUuser2.optChg = vbUnchecked
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Set AdmnUsrpw = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtPas = String(15, " ")
   txtVpw = String(15, " ")
   
End Sub

Private Sub FillUser()
   lblUser = SecPw.UserLcName
   lblName = Secure.UserName
   sNewPass = GetSecPassword(SecPw.PassWord)
   sVerify = sNewPass
   txtPas = sNewPass
   txtVpw = sVerify
   Exit Sub
   
   DiaErr1:
   sProcName = "filluser"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub txtPas_LostFocus()
   txtPas = CheckLen(txtPas, 15)
   If txtPas = "" Then txtPas = String(15, " ")
   cmdChg.Enabled = True
   
End Sub


Private Sub txtVpw_LostFocus()
   txtVpw = CheckLen(txtVpw, 15)
   If txtVpw = "" Then txtVpw = String(15, " ")
   
End Sub
