VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form SadmSLe02a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Regions"
   ClientHeight    =   3030
   ClientLeft      =   1200
   ClientTop       =   855
   ClientWidth     =   5625
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   HelpContextID   =   1202
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   5625
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "SadmSLe02a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.Timer Timer1 
      Left            =   240
      Top             =   1800
   End
   Begin VB.TextBox txtDsc 
      Height          =   285
      Left            =   1360
      TabIndex        =   1
      Top             =   960
      Width           =   3495
   End
   Begin VB.CheckBox optUse 
      Alignment       =   1  'Right Justify
      Caption         =   "Force Regions"
      Enabled         =   0   'False
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   2400
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.TextBox txtCmt 
      Height          =   975
      Left            =   1360
      MultiLine       =   -1  'True
      TabIndex        =   2
      Tag             =   "9"
      Top             =   1320
      Width           =   3475
   End
   Begin VB.ComboBox cmbReg 
      Height          =   315
      Left            =   1360
      Sorted          =   -1  'True
      TabIndex        =   0
      ToolTipText     =   "Enter/Revise A Region (2 char)"
      Top             =   600
      Width           =   660
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   4680
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
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
      FormDesignHeight=   3030
      FormDesignWidth =   5625
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Forces Use Of Fixed Regions In All Sections)"
      Height          =   255
      Index           =   3
      Left            =   1680
      TabIndex        =   8
      Top             =   2445
      Width           =   4095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Comment"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Region"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "SadmSLe02a"
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
Dim RdoRgn As ADODB.Recordset

Dim bOnLoad As Byte
Dim bGoodReg As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbReg_Click()
   bGoodReg = GetRegion()
   
End Sub


Private Sub cmbReg_LostFocus()
   cmbReg = CheckLen(cmbReg, 2)
   If Len(cmbReg) Then
      cmbReg = Compress(cmbReg)
      bGoodReg = GetRegion()
      If bGoodReg = 0 Then AddRegion
   Else
      bGoodReg = 0
   End If
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   cmbReg = ""
   
End Sub


Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 1202
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub


Private Sub Form_Activate()
   If bOnLoad Then
      FillRegions
      bOnLoad = 0
      If cmbReg.ListCount > 0 Then
         cmbReg = cmbReg.List(0)
         bGoodReg = GetRegion()
      End If
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   GetCurrentSelections
   
   sSql = "SELECT TOP 1 REGREF,REGDESC,REGCOMT FROM " _
          & "CregTable WHERE REGREF= ? "
          
   Set AdoQry = New ADODB.Command
   AdoQry.CommandText = sSql
   
   Set AdoParameter = New ADODB.Parameter
   AdoParameter.Type = adChar
   AdoParameter.Size = 4
   
   AdoQry.Parameters.Append AdoParameter
          

   bOnLoad = 1
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   cUR.CurrentRegion = cmbReg
   SaveCurrentSelections
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set AdoParameter = Nothing
   Set AdoQry = Nothing
   Set SadmSLe02a = Nothing
   
End Sub


Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub optUse_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub txtCmt_LostFocus()
   txtCmt = CheckLen(txtCmt, 255)
   txtCmt = StrCase(txtCmt, ES_FIRSTWORD)
   On Error Resume Next
   If bGoodReg Then
      RdoRgn!REGCOMT = "" & txtCmt
      RdoRgn.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtDsc_LostFocus()
   txtDsc = CheckLen(txtDsc, 40)
   txtDsc = StrCase(txtDsc)
   On Error Resume Next
   If bGoodReg Then
      RdoRgn!REGDESC = "" & txtDsc
      RdoRgn.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub



Private Function GetRegion() As Byte
   Dim sRegion As String
   sRegion = cmbReg
   On Error GoTo DiaErr1
   'RdoQry(0) = sRegion
   AdoQry.Parameters(0).value = sRegion
   
   bSqlRows = clsADOCon.GetQuerySet(RdoRgn, AdoQry, ES_KEYSET, True)
   If bSqlRows Then
      With RdoRgn
         cmbReg = "" & Trim(!REGREF)
         txtDsc = "" & Trim(!REGDESC)
         txtCmt = "" & Trim(!REGCOMT)
      End With
      GetRegion = 1
   Else
      txtDsc = ""
      txtCmt = ""
      GetRegion = 0
   End If
   Exit Function
   
DiaErr1:
   sProcName = "getregion"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub AddRegion()
   Dim bResponse As Byte
   Dim sMsg As String
   Dim sRegion As String
   
   sRegion = cmbReg
   bResponse = IllegalCharacters(cmbReg)
   If bResponse > 0 Then
      MsgBox "The Region Contains An Illegal " & Chr$(bResponse) & ".", _
         vbExclamation, Caption
      Exit Sub
   End If
   sMsg = sRegion & " Wasn't Found. Add The Region?"
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      On Error GoTo DiaErr1
      sSql = "INSERT INTO CregTable (REGREF) " _
             & "VALUES('" & sRegion & "')"
      clsADOCon.ExecuteSQL sSql
      If clsADOCon.RowsAffected Then
         SysMsg "Region Added.", True
         cmbReg = sRegion
         AddComboStr cmbReg.hwnd, sRegion
         bGoodReg = GetRegion()
         On Error Resume Next
         txtDsc.SetFocus
      Else
         MsgBox "Couldn't The Add Region.", vbExclamation, Caption
      End If
   Else
      CancelTrans
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "addregion"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub
