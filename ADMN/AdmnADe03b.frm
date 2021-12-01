VERSION 5.00
Begin VB.Form AdmnADe03b 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Class"
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6120
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   Icon            =   "AdmnADe03b.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   6120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtTtt 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Tag             =   "2"
      ToolTipText     =   "Used For The Pop Tool Tip (40)"
      Top             =   960
      Width           =   3255
   End
   Begin VB.TextBox txtCls 
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Tag             =   "2"
      ToolTipText     =   "New Class Id (20 And Min 3)"
      Top             =   600
      Width           =   2295
   End
   Begin VB.CommandButton cmdClass 
      Caption         =   "&Add"
      Enabled         =   0   'False
      Height          =   315
      Left            =   5160
      TabIndex        =   3
      ToolTipText     =   "Add Some Class"
      Top             =   480
      Width           =   875
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5160
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin VB.Label lblListIndex 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tool Tip Text"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Comment Class"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label z1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Width "
      Height          =   255
      Index           =   3
      Left            =   3700
      TabIndex        =   4
      Top             =   600
      Width           =   700
   End
End
Attribute VB_Name = "AdmnADe03b"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bCancel As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   bCancel = 1
   
End Sub


Private Sub cmdClass_Click()
   Dim bResponse As Byte
   If Len(txtTtt) = 0 Then
      MsgBox "Please Enter A Short Description For The Class.", _
         vbInformation, Caption
      On Error Resume Next
      txtTtt.SetFocus
      Exit Sub
   Else
      bResponse = MsgBox("Are You Ready To Add The New Class?", _
                  ES_YESQUESTION, Caption)
      If bResponse = vbYes Then
        AddClass
      Else
        CancelTrans
      End If
   End If
   
End Sub

Private Sub Form_Load()
   AlwaysOnTop hwnd, True
   AdmnADe03a.Enabled = False
   If iBarOnTop Then
      Move MDISect.Left + 400, MDISect.Top + 2600
   Else
      Move MDISect.Left + 2400, MDISect.Top + 2000
   End If
   FormatControls
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   AdmnADe03a.Enabled = True
   AlwaysOnTop hwnd, False
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Set AdmnADe03b = Nothing
   
End Sub


Private Sub txtCls_Change()
   txtCls = CheckLen(txtCls, 20)
   txtCls = StrCase(txtCls)
   If bCancel Then Exit Sub
   If Len(txtCls) > 3 Then
      cmdClass.Enabled = True
   Else
      MsgBox "Please Make Your Class At Least (3) Characters.", _
         vbInformation, Caption
   End If

End Sub

Private Sub txtCls_LostFocus()
   txtCls = CheckLen(txtCls, 20)
   txtCls = StrCase(txtCls)
   If bCancel Then Exit Sub
   If Len(txtCls) > 3 Then
      cmdClass.Enabled = True
   Else
      MsgBox "Please Make Your Class At Least (3) Characters.", _
         vbInformation, Caption
   End If
   
End Sub


Private Sub txtTtt_LostFocus()
   txtTtt = CheckLen(txtTtt, 40)
   txtTtt = StrCase(txtTtt)
   
End Sub



Private Sub AddClass()
   Dim RdoCls As ADODB.Recordset
   On Error Resume Next
   clsADOCon.ADOErrNum = 0
   sSql = "SELECT * FROM StchTable WHERE COMMENT_CLASS=''"
   
   Set RdoCls = clsADOCon.GetRecordSet(sSql, ES_DYNAMIC)
   With RdoCls
      .AddNew
      !COMMENT_CLASS = Trim(txtCls)
      !COMMENT_LISTINDEX = Val(lblListIndex)
      !COMMENT_TOOLTIP = Trim(txtTtt)
      !COMMENT_WIDTH = Val(Right(z1(3), 2))
      !COMMENT_USER = 1
      .Update
      If clsADOCon.ADOErrNum = 0 Then
         SysMsg "Class Was Added.", True
         Sleep 1000
         AdmnADe03a.lblListIndex = 8
         AdmnADe03a.cmbCls = txtCls
         Unload Me
      Else
         MsgBox "Couldn't Add Class.", _
            vbInformation, Caption
         AdmnADe03a.cmbCls = txtCls
         Unload Me
      End If
      ClearResultSet RdoCls
   End With
   Set RdoCls = Nothing
   
End Sub
