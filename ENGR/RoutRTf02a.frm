VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form RoutRTf02a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Copy A Routing "
   ClientHeight    =   2865
   ClientLeft      =   2430
   ClientTop       =   1515
   ClientWidth     =   5655
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2865
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "RoutRTf02a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdVew 
      DownPicture     =   "RoutRTf02a.frx":07AE
      Height          =   320
      Left            =   4680
      Picture         =   "RoutRTf02a.frx":0C88
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "Parts Assigned To This Routing"
      Top             =   1200
      Width           =   350
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "C&opy"
      Height          =   315
      Left            =   4680
      TabIndex        =   5
      ToolTipText     =   "Copy The Existing To The New Routing"
      Top             =   1920
      Width           =   875
   End
   Begin VB.TextBox txtNew 
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Tag             =   "3"
      Top             =   1980
      Width           =   3075
   End
   Begin VB.ComboBox cmbRte 
      Height          =   315
      Left            =   1320
      TabIndex        =   0
      Tag             =   "3"
      Top             =   1200
      Width           =   3345
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   4680
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   120
      Top             =   2640
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2865
      FormDesignWidth =   5655
   End
   Begin MSComctlLib.ProgressBar prg1 
      Height          =   300
      Left            =   1320
      TabIndex        =   9
      Top             =   2400
      Width           =   3012
      _ExtentX        =   5318
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label txtDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1320
      TabIndex        =   7
      Top             =   1560
      Width           =   3075
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "New Number"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   4
      Top             =   1980
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Old Routing"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   1170
      Width           =   1215
   End
End
Attribute VB_Name = "RoutRTf02a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'12/3/05 Restructured copy and add bShow to FormUnload
'10/3/06 CopyRouting completely revamped with SQL Commands
Option Explicit
Dim bGoodOld As Byte
Dim bGoodNew As Byte
Dim bOnLoad As Byte
Dim bShow As Byte

Dim sOldRout As String
Dim sNewRout As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd


Private Sub cmbRte_Click()
   bGoodOld = GetRout(True)
   
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
      OpenHelpContext 3151
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub cmdNew_Click()
   bGoodNew = GetRout(False)
   If bGoodNew Then CopyRouting
   
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
      If cmbRte.ListCount > 0 Then bGoodOld = GetRout(True)
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST
   FormatControls
   bOnLoad = 1
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If bGoodNew Then
      sCurrRout = txtNew
      SaveSetting "Esi2000", "EsiEngr", "CurrentRouting", Trim(sCurrRout)
   Else
      sCurrRout = ""
   End If
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   If bShow = 0 Then FormUnload
   Set RoutRTf02a = Nothing
   
End Sub




Private Sub txtNew_LostFocus()
   txtNew = CheckLen(txtNew, 30)
   If Len(txtNew) = 0 Then Exit Sub
   sOldRout = Compress(cmbRte)
   sNewRout = Compress(txtNew)
   If sOldRout = sNewRout Then
      MsgBox "New Number Is The Same.", vbInformation, Caption
      Exit Sub
   End If
   
End Sub



Private Function GetRout(bFromCombo As Byte) As Byte
   Dim RdoRte As ADODB.Recordset
   Dim sRout As String
   If bFromCombo Then
      sRout = Compress(cmbRte)
   Else
      sRout = Compress(txtNew)
   End If
   GetRout = 0
   On Error GoTo DiaErr1
   sSql = "Qry_GetToolRout '" & sRout & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRte, ES_STATIC)
   If bSqlRows Then
      If bFromCombo Then
         GetRout = 0
         cmbRte = "" & Trim(RdoRte!RTNUM)
         txtDsc = "" & Trim(RdoRte!RTDESC)
         ClearResultSet RdoRte
      Else
         MsgBox "That Routing Is Already Recorded.", vbInformation, Caption
         GetRout = 0
      End If
   Else
      If bFromCombo Then
         MsgBox "That Routing Wasn't Found.", vbExclamation, Caption
         GetRout = 0
      Else
         GetRout = 1
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

Private Sub CopyRouting()
   Dim bResponse As Byte
   Dim RdoCpy As ADODB.Recordset
   Dim RdoRte As ADODB.Recordset
   
   bResponse = MsgBox("Copy Routing " & Trim(cmbRte) & " To " _
               & Trim(txtNew) & ".", ES_YESQUESTION, Caption)
   If bResponse = vbNo Then
      On Error Resume Next
      cmdCan.SetFocus
      Width = Width + 10
      Exit Sub
   End If
   sNewRout = Compress(txtNew)
   sOldRout = Compress(cmbRte)
   
   MouseCursor 11
   cmdCan.Enabled = False
   prg1.Visible = True
   prg1.Value = 10
   
   On Error Resume Next
   MouseCursor 13
   'In case the temp table remains
   sSql = "DROP TABLE #Rthd"
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   
   sSql = "DROP TABLE #Rtop"
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   
   'New Header
   Err.Clear
   clsADOCon.ADOErrNum = 0
   clsADOCon.BeginTrans
   
   sSql = "SELECT * INTO #Rthd from RthdTable where RTREF='" & sOldRout & "' "
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   prg1.Value = 30
   Sleep 200
'bbs changed on 3/14/2016 to not remove revision notes
'   sSql = "UPDATE #Rthd SET RTREF='" & sNewRout & "',RTNUM='" & txtNew & "'," _
'              & "RTDATE='" & Format(Now, "mm/dd/yy") & "', RTREVNOTES = '', RTAPPDATE = NULL," _
'          & "RTAPPBY = ''"
'          '& "RTDATE='" & Format(Now, "mm/dd/yy") & "'"
   sSql = "UPDATE #Rthd SET RTREF='" & sNewRout & "',RTNUM='" & txtNew & "'," _
              & "RTDATE='" & Format(Now, "mm/dd/yy") & "', RTAPPDATE = NULL," _
          & "RTAPPBY = ''"
          '& "RTDATE='" & Format(Now, "mm/dd/yy") & "'"
   
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   
   sSql = "INSERT INTO RthdTable SELECT * FROM #Rthd"
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   
   prg1.Value = 50
   Sleep 200
   
   'New Operations
   sSql = "SELECT * INTO #Rtop from RtopTable where OPREF='" & sOldRout & "' "
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   prg1.Value = 70
   Sleep 200
   
   sSql = "UPDATE #Rtop SET OPREF='" & sNewRout & "'"
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   
   sSql = "INSERT INTO RtopTable SELECT * FROM #Rtop"
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   prg1.Value = 100
   MouseCursor 0
   If clsADOCon.ADOErrNum = 0 Then
      clsADOCon.CommitTrans
      MsgBox "Routing Successfully Copied.", _
         vbInformation, Caption
      bShow = 1
      sCurrRout = txtNew
      RoutRTe01a.cmbRte = txtNew
      RoutRTe01a.Show
      Unload Me
   Else
      MsgBox "Could Not Copy The Routing.", _
         vbInformation, Caption
      prg1.Visible = False
      ' TODO: Added not There
      clsADOCon.RollbackTrans
      clsADOCon.ADOErrNum = 0
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "copyrouting"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub
