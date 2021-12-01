VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form RoutRTe04a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Revise A Routing Number"
   ClientHeight    =   2310
   ClientLeft      =   2430
   ClientTop       =   1515
   ClientWidth     =   5010
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2310
   ScaleWidth      =   5010
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "RoutRTe04a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdUpd 
      Caption         =   "&Apply"
      Height          =   315
      Left            =   4080
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Change The Old Routing To The New Routing Number (ID)"
      Top             =   600
      Width           =   875
   End
   Begin VB.TextBox txtNew 
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Top             =   1500
      Width           =   3075
   End
   Begin VB.ComboBox cmbRte 
      Height          =   315
      Left            =   1320
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   1080
      Width           =   3345
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   4080
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   4560
      Top             =   1800
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2310
      FormDesignWidth =   5010
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "New Number"
      Height          =   285
      Index           =   1
      Left            =   180
      TabIndex        =   4
      Top             =   1500
      Width           =   2400
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Old Routing"
      Height          =   285
      Index           =   0
      Left            =   180
      TabIndex        =   3
      Top             =   1050
      Width           =   2400
   End
End
Attribute VB_Name = "RoutRTe04a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'1/26/05 Added Illegal Characters
Option Explicit
Dim bGoodOld As Byte
Dim bGoodNew As Byte
Dim bOnLoad As Byte

Dim sOldRout As String
Dim sNewRout As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub cmbRte_LostFocus()
   cmbRte = CheckLen(cmbRte, 30)
   If Len(cmbRte) = 0 Then
      bGoodOld = False
      Exit Sub
   Else
      bGoodOld = GetRout(True)
   End If
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub


Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   cmdCan_Click
   
End Sub


Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 3104
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub


Private Sub cmdUpd_Click()
   txtNew = CheckLen(txtNew, 30)
   If Len(txtNew) = 0 Then
      MsgBox "Requires A Valid Routing Number.", vbInformation, Caption
   Else
      sOldRout = Compress(cmbRte)
      sNewRout = Compress(txtNew)
      If sOldRout = sNewRout Then
         MsgBox "New Number Is The Same.", vbInformation, Caption
      Else
         bGoodNew = GetRout(False)
         If bGoodNew Then ReviseRouting
      End If
   End If
   
End Sub


Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      bOnLoad = 0
      FillRoutings
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
   Set RoutRTe04a = Nothing
   
End Sub




Private Sub txtNew_LostFocus()
   txtNew = CheckLen(txtNew, 30)
   
End Sub



Private Function GetRout(bFromCombo As Byte) As Byte
   Dim RdoRte As ADODB.Recordset
   Dim sRout As String
   If bFromCombo Then
      sRout = Compress(cmbRte)
   Else
      sRout = Compress(txtNew)
   End If
   GetRout = False
   On Error GoTo DiaErr1
   sSql = "Qry_GetRoutingBasics '" & sRout & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRte, ES_STATIC)
   If bSqlRows Then
      If bFromCombo Then
         GetRout = True
         cmbRte = "" & Trim(RdoRte!RTNUM)
      Else
         MsgBox "That Routing Is Already Recorded.", vbInformation, Caption
         GetRout = False
      End If
   Else
      If bFromCombo Then
         MsgBox "Routing Wasn't Found.", vbExclamation, Caption
         GetRout = False
      Else
         GetRout = True
      End If
   End If
   Set RdoRte = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getrout"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub ReviseRouting()
   Dim bResponse As Byte
   bResponse = IllegalCharacters(txtNew)
   If bResponse > 0 Then
      MsgBox "Routing Number Contains An Illegal " & Chr$(bResponse) & ".", _
         vbExclamation, Caption
      Exit Sub
   End If
   
   'determine whether routing is in use
   sNewRout = Compress(txtNew)
   sOldRout = Compress(cmbRte)
   Dim rdo As ADODB.Recordset
   sSql = "select * from RthdTable where RTREF = '" & sNewRout & "'"
   If clsADOCon.GetDataSet(sSql, rdo) Then
      MsgBox "That routing number is already in use"
      Exit Sub
   End If
   
   bResponse = MsgBox("Are You Sure That You Want To Change The Routing Number?", ES_NOQUESTION, Caption)
   If bResponse = vbNo Then
      'On Error Resume Next
      cmdCan.SetFocus
      Width = Width + 10
      Exit Sub
   End If
   
   MouseCursor ccHourglass
   cmdCan.Enabled = False
   
   On Error GoTo RrevnRr1
   
'  can only do one update at a time
'   sSql = "UPDATE ComnTable SET RTEPART1='" & txtNew & "' WHERE RTEPART1='" & sOldRout & "';" _
'          & "UPDATE ComnTable SET RTEPART2='" & txtNew & "' WHERE RTEPART2='" & sOldRout & "';" _
'          & "UPDATE ComnTable SET RTEPART3='" & txtNew & "' WHERE RTEPART3='" & sOldRout & "';" _
'          & "UPDATE ComnTable SET RTEPART4='" & txtNew & "' WHERE RTEPART4='" & sOldRout & "';" _
'          & "UPDATE ComnTable SET RTEPART5='" & txtNew & "' WHERE RTEPART5='" & sOldRout & "';" _
'          & "UPDATE ComnTable SET RTEPART6='" & txtNew & "' WHERE RTEPART6='" & sOldRout & "';" _
'          & "UPDATE ComnTable SET RTEPART8='" & txtNew & "' WHERE RTEPART8='" & sOldRout & "';" _
'          & "UPDATE RthdTable SET RTREF='" & sNewRout & "',RTNUM='" & txtNew & "' WHERE RTREF='" & sOldRout & "';" _
'          & "UPDATE RtopTable SET OPREF='" & sNewRout & "' WHERE OPREF='" & sOldRout & "';" _
'          & "UPDATE RtpcTable SET OPREF='" & sNewRout & "' WHERE OPREF='" & sOldRout & "';" _
'          & "UPDATE PartTable SET PAROUTING='" & sNewRout & "' WHERE PAROUTING='" & sOldRout & "'"
   
   clsADOCon.BeginTrans
   
   sSql = "INSERT INTO RthdTable(  RTREF,RTNUM,RTBY,RTDATE, RTAPPBY, RTAPPDATE, RTREV," _
            & " RTDESC, RTQUEUEHRS ,RTMOVEHRS ,RTSETUPHRS ,RTREVNOTES)" _
         & " SELECT '" & sNewRout & "','" & txtNew & "',RTBY,RTDATE, RTAPPBY, RTAPPDATE, RTREV," _
            & " RTDESC, RTQUEUEHRS ,RTMOVEHRS ,RTSETUPHRS ,RTREVNOTES FROM RthdTable " _
         & " WHERE RTREF ='" & sOldRout & "'"
   clsADOCon.ExecuteSQL sSql ', rdExecDirect
   
   
   sSql = "UPDATE ComnTable SET RTEPART1='" & sNewRout & "' WHERE RTEPART1='" & sOldRout & "'"
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   sSql = "UPDATE ComnTable SET RTEPART2='" & sNewRout & "' WHERE RTEPART2='" & sOldRout & "'"
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   sSql = "UPDATE ComnTable SET RTEPART3='" & sNewRout & "' WHERE RTEPART3='" & sOldRout & "'"
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   sSql = "UPDATE ComnTable SET RTEPART4='" & sNewRout & "' WHERE RTEPART4='" & sOldRout & "'"
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   sSql = "UPDATE ComnTable SET RTEPART5='" & sNewRout & "' WHERE RTEPART5='" & sOldRout & "'"
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   sSql = "UPDATE ComnTable SET RTEPART6='" & sNewRout & "' WHERE RTEPART6='" & sOldRout & "'"
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   sSql = "UPDATE ComnTable SET RTEPART8='" & sNewRout & "' WHERE RTEPART8='" & sOldRout & "'"
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   
   sSql = "UPDATE RtpcTable SET OPREF='" & sNewRout & "' WHERE OPREF='" & sOldRout & "'"
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   
   sSql = "UPDATE PartTable SET PAROUTING='" & sNewRout & "' WHERE PAROUTING='" & sOldRout & "'"
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   
   'sSql = "UPDATE RthdTable SET RTREF='" & sNewRout & "',RTNUM='" & txtNew & "' WHERE RTREF='" & sOldRout & "'"
   'RdoCon.Execute sSql, rdExecDirect
   
   sSql = "UPDATE RtopTable SET OPREF='" & sNewRout & "' WHERE OPREF='" & sOldRout & "'"
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   
   sSql = "DELETE FROM RthdTable WHERE RTREF='" & sOldRout & "'"
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   
   
   clsADOCon.CommitTrans
   MouseCursor 0
   FillRoutings
   MouseCursor ccDefault
   MsgBox "Routing Number Changed.", vbInformation, Caption
   cmbRte = txtNew
   txtNew = ""
   cmdCan.Enabled = True
   Exit Sub
   
'RrevnRr1:
'   sProcName = "ReviseRouting"
'   CurrError.Number = Err.Number
'   CurrError.Description = Err.Description
'   Resume RrevnRr2
'RrevnRr2:
'   On Error Resume Next
'   clsADOCon.RollbackTrans
'   cmdCan.Enabled = True
'   MsgBox CurrError.Description, vbExclamation, Caption
'   DoModuleErrors Me
   
RrevnRr1:
   clsADOCon.RollbackTrans
   MouseCursor ccDefault
   sProcName = "ReviseRouting"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   cmdCan.Enabled = True
   MsgBox CurrError.Description, vbExclamation, Caption
   DoModuleErrors Me
   
End Sub
