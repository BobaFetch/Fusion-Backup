VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form diaHcode 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Time Type Codes"
   ClientHeight    =   3090
   ClientLeft      =   1200
   ClientTop       =   855
   ClientWidth     =   5520
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtAdd 
      Height          =   285
      Left            =   1440
      TabIndex        =   6
      ToolTipText     =   "Multiply Rate By This Number (Never 0)"
      Top             =   2400
      Width           =   615
   End
   Begin VB.TextBox txtSeq 
      Height          =   285
      Left            =   1440
      TabIndex        =   2
      ToolTipText     =   "ComboBox Sort Sequence"
      Top             =   1320
      Width           =   375
   End
   Begin VB.Frame z2 
      Height          =   615
      Left            =   1440
      TabIndex        =   13
      Top             =   1680
      Width           =   3975
      Begin VB.OptionButton optDbl 
         Caption         =   "Double Time"
         Height          =   195
         Left            =   2640
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optOvr 
         Caption         =   "Over Time"
         Height          =   195
         Left            =   1340
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optReg 
         Caption         =   "Regular"
         Height          =   195
         Left            =   100
         TabIndex        =   3
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
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
      FormDesignHeight=   3090
      FormDesignWidth =   5520
   End
   Begin VB.TextBox txtDsc 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Top             =   960
      Width           =   3495
   End
   Begin VB.ComboBox cmbCde 
      Height          =   315
      Left            =   1440
      Sorted          =   -1  'True
      TabIndex        =   0
      ToolTipText     =   "Enter/Revise A Division (2 char)"
      Top             =   600
      Width           =   660
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   4560
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   90
      Width           =   875
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   10
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
      PictureUp       =   "diaHcode.frx":0000
      PictureDn       =   "diaHcode.frx":0146
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Ex: Reg = 1.0, OT = 1.5)"
      Height          =   255
      Index           =   5
      Left            =   2160
      TabIndex        =   15
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Rate Multiplier"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   14
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Time Charge"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   12
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sort Sequence"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   11
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Time Type Code"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "diaHcode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of            ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
Option Explicit
Dim AdoCde As ADODB.Recordset

Dim bOnLoad As Byte
Dim bGoodCode As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd


Private Sub cmbCde_Click()
   bGoodCode = GetTimeCode()
   
End Sub

Private Sub cmbCde_LostFocus()
   cmbCde = CheckLen(cmbCde, 2)
   If Len(cmbCde) Then
      bGoodCode = GetTimeCode()
      If Not bGoodCode Then AddTimeCode
   Else
      bGoodCode = False
   End If
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
   cmbCde = ""
   
End Sub


Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext "hs1504"
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub


Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      FillCombo
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   
   AddDefaults
   bOnLoad = 1
   Show
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set AdoCde = Nothing
   Set diaHcode = Nothing
   
End Sub




Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub




Private Sub txtAdd_LostFocus()
   txtAdd = CheckLen(txtAdd, 5)
   txtAdd = Format(Abs(Val(txtAdd)), "0.000")
   If Val(txtAdd) = 0 Then
      'Beep
      txtAdd = Format(1, "0.000")
   End If
   If bGoodCode Then
      'AdoCde.Edit
      AdoCde!TYPEADDER = Val(txtAdd)
      AdoCde.Update
      If Err > 0 Then ValidateEdit Me
   End If
   
End Sub


Private Sub txtDsc_LostFocus()
   txtDsc = CheckLen(txtDsc, 40)
   txtDsc = StrCase(txtDsc)
   On Error Resume Next
   If bGoodCode Then
      'AdoCde.Edit
      AdoCde!TYPEDESC = "" & txtDsc
      AdoCde.Update
      If Err > 0 Then ValidateEdit Me
   End If
   
End Sub



Private Function GetTimeCode() As Byte
   Dim sTimeCode As String
   Dim sType As String
   
   sTimeCode = Compress(cmbCde)
   On Error GoTo DiaErr1

   
   sSql = "SELECT TOP 1 TYPECODE,TYPEDESC," _
          & "TYPESEQ,TYPETYPE,TYPEADDER FROM " _
          & "TmcdTable WHERE TYPECODE= '" & sTimeCode & "' "
    bSqlRows = clsADOCon.GetDataSet(sSql, AdoCde, ES_DYNAMIC)
   If bSqlRows Then
      With AdoCde
         cmbCde = "" & Trim(!typeCode)
         txtDsc = "" & Trim(!TYPEDESC)
         txtSeq = Format(0 + !TYPESEQ, "#0")
         txtAdd = Format(0 + !TYPEADDER, "#0.000")
         sType = "" & Trim(!TYPETYPE)
         Select Case sType
            Case "R"
               optReg.Value = True
            Case "O"
               optOvr.Value = True
            Case "D"
               optDbl.Value = True
         End Select
         sType = cmbCde
         'If sType = "RT" Then txtAdd.Enabled = False _
         'Else txtAdd.Enabled = True
         txtAdd.Enabled = True
         If sType = "RT" Or sType = "OT" Or sType = "DT" Then
            optReg.Enabled = False
            optOvr.Enabled = False
            optDbl.Enabled = False
         Else
            optReg.Enabled = True
            optOvr.Enabled = True
            optDbl.Enabled = True
            txtAdd.Enabled = True
         End If
         .Cancel
      End With
      GetTimeCode = True
   Else
      txtDsc = ""
      GetTimeCode = False
   End If
   'Set AdoCde = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "gettimeco"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub AddTimeCode()
   Dim bResponse As Byte
   Dim sMsg As String
   Dim sTimeCode As String
   
   sTimeCode = Compress(cmbCde)
   sMsg = sTimeCode & " Wasn't Found. Add The Time Code?"
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      On Error GoTo DiaErr1
      sSql = "INSERT INTO TmcdTable (TYPECODE) " _
             & "VALUES('" & sTimeCode & "')"
      clsADOCon.ExecuteSql sSql
      If clsADOCon.RowsAffected Then
         SysMsg "Time Code Added.", True
         cmbCde = sTimeCode
         AddComboStr cmbCde.hwnd, sTimeCode
         bGoodCode = GetTimeCode()
         On Error Resume Next
         txtDsc.SetFocus
      Else
         MsgBox "Couldn't The Add Time Code.", vbExclamation, Caption
      End If
   Else
      CancelTrans
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "addtimeco"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub FillCombo()
   On Error GoTo DiaErr1
   sSql = "SELECT TYPECODE FROM TmcdTable "
   LoadComboBox cmbCde, -1
   If cmbCde.ListCount > 0 Then
      cmbCde = cmbCde.List(0)
      bGoodCode = GetTimeCode
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub AddDefaults()
   'For new systems, add defaults if they don't exist
   On Error GoTo DiaErr1
   sSql = "SELECT TYPECODE FROM TmcdTable"
   bSqlRows = clsADOCon.GetDataSet(sSql, AdoCde, ES_FORWARD)
   If Not bSqlRows Then
      sSql = "INSERT INTO TmcdTable (TYPECODE,TYPEDESC,TYPESEQ,TYPETYPE) " _
             & "VALUES('RT','Regular Time',0,'R')"
      clsADOCon.ExecuteSql sSql
      
      sSql = "INSERT INTO TmcdTable (TYPECODE,TYPEDESC,TYPESEQ,TYPETYPE) " _
             & "VALUES('OT','Overtime',1,'O')"
      clsADOCon.ExecuteSql sSql
      
      sSql = "INSERT INTO TmcdTable (TYPECODE,TYPEDESC,TYPESEQ,TYPETYPE) " _
             & "VALUES('DT','Double Time',2,'D')"
      clsADOCon.ExecuteSql sSql
   End If
   Exit Sub
DiaErr1:
   'just bail
   On Error GoTo 0
   
End Sub

Private Sub txtSeq_LostFocus()
   txtSeq = CheckLen(txtSeq, 2)
   txtSeq = Format(Abs(Val(txtSeq)), "#0")
   If bGoodCode Then
      'RdoCde.Edit
      AdoCde!TYPESEQ = Val(txtSeq)
      AdoCde.Update
      If Err > 0 Then ValidateEdit Me
   End If
   
End Sub
