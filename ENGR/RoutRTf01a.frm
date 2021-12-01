VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form RoutRTf01a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Delete A Routing"
   ClientHeight    =   2145
   ClientLeft      =   2430
   ClientTop       =   1515
   ClientWidth     =   6480
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2145
   ScaleWidth      =   6480
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "RoutRTf01a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdVew 
      DownPicture     =   "RoutRTf01a.frx":07AE
      Height          =   320
      Left            =   4940
      Picture         =   "RoutRTf01a.frx":0C88
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Parts Assigned To This Routing"
      Top             =   820
      Width           =   360
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "&Delete"
      Height          =   315
      Left            =   5520
      TabIndex        =   3
      ToolTipText     =   "Delete This Routing"
      Top             =   480
      Width           =   875
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5520
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin VB.ComboBox cmbRte 
      Height          =   315
      Left            =   1560
      TabIndex        =   0
      Top             =   810
      Width           =   3345
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6000
      Top             =   1320
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2145
      FormDesignWidth =   6480
   End
   Begin VB.Label txtDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1560
      TabIndex        =   5
      Top             =   1200
      Width           =   3075
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Routing Number"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   810
      Width           =   1545
   End
End
Attribute VB_Name = "RoutRTf01a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'7/20/06 Add reference to RtpcTable
Option Explicit
Dim bGoodOld As Byte
Dim bOnLoad As Byte

Dim sOldRout As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbRte_LostFocus()
   cmbRte = CheckLen(cmbRte, 30)
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub


Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   cmdCan_Click
   
End Sub


Private Sub cmdDel_Click()
   bGoodOld = GetRout()
   If bGoodOld Then ReviseRouting
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 3150
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub cmdVew_Click()
   If cmdVew Then
      RteTree.Show
      cmdVew = False
   End If
   
End Sub

Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      bOnLoad = 0
      FillRoutings
      If cmbRte.ListCount > 0 Then bGoodOld = GetRout()
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
   On Error Resume Next
   FormUnload
   Set RoutRTf01a = Nothing
   
End Sub












Private Function GetRout() As Byte
   Dim RdoRte As ADODB.Recordset
   Dim sRout As String
   sRout = Compress(cmbRte)
   GetRout = False
   On Error GoTo DiaErr1
   sSql = "Qry_GetToolRout '" & sRout & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRte)
   If bSqlRows Then
      With RdoRte
         GetRout = True
         cmbRte = "" & Trim(RdoRte!RTNUM)
         txtDsc = "" & Trim(RdoRte!RTDESC)
         ClearResultSet RdoRte
      End With
   Else
      GetRout = False
   End If
   If Not GetRout Then MsgBox "Couldn't Find The Routing.", vbInformation, Caption
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
   
   bResponse = MsgBox("Are You Sure That You Want To Delete The Routing.", ES_NOQUESTION, Caption)
   If bResponse = vbNo Then
      On Error Resume Next
      cmdCan.SetFocus
      Width = Width + 10
      Exit Sub
   End If
   sOldRout = Compress(cmbRte)
   
   MouseCursor 11
   cmdCan.Enabled = False
   On Error GoTo DiaErr1
   clsADOCon.ADOErrNum = 0
   clsADOCon.BeginTrans
    sSql = "UPDATE ComnTable SET RTEPART1='' WHERE RTEPART1='" & sOldRout & "';" _
          & "UPDATE ComnTable SET RTEPART2='' WHERE RTEPART2='" & sOldRout & "';" _
          & "UPDATE ComnTable SET RTEPART3='' WHERE RTEPART3='" & sOldRout & "';" _
          & "UPDATE ComnTable SET RTEPART4='' WHERE RTEPART4='" & sOldRout & "';" _
          & "UPDATE ComnTable SET RTEPART5='' WHERE RTEPART5='" & sOldRout & "';" _
          & "UPDATE ComnTable SET RTEPART6='' WHERE RTEPART6='" & sOldRout & "';" _
          & "UPDATE ComnTable SET RTEPART8='' WHERE RTEPART8='" & sOldRout & "';"
   
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
          
      
      sSql = "DELETE FROM RtopTable WHERE OPREF='" & sOldRout & "';" _
          & "DELETE FROM RtpcTable WHERE OPREF='" & sOldRout & "';" _
          & "UPDATE PartTable SET PAROUTING='' WHERE PAROUTING='" & sOldRout & "'"
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   
   sSql = "DELETE FROM RthdTable WHERE RTREF='" & sOldRout & "';"
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   
   If clsADOCon.ADOErrNum = 0 Then
      clsADOCon.CommitTrans
      MsgBox "Routing Deleted and Assignments Reset.", vbInformation, Caption
      Unload Me
   Else
      clsADOCon.RollbackTrans
      MsgBox "Could Delete Routing and Reset Assignments.", _
         vbExclamation, Caption
   End If
   
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   On Error Resume Next
   sProcName = "reviserouting"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   clsADOCon.RollbackTrans
   cmdCan.Enabled = True
   DoModuleErrors Me
   
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub
