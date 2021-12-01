VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form RoutRTf05a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reorganize Operations"
   ClientHeight    =   4260
   ClientLeft      =   3015
   ClientTop       =   1560
   ClientWidth     =   5670
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4260
   ScaleWidth      =   5670
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdUp 
      DisabledPicture =   "RoutRTf05a.frx":0000
      Height          =   372
      Left            =   4920
      MaskColor       =   &H00000000&
      Picture         =   "RoutRTf05a.frx":04F2
      Style           =   1  'Graphical
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   3240
      Width           =   400
   End
   Begin VB.CommandButton cmdDn 
      DisabledPicture =   "RoutRTf05a.frx":09E4
      Height          =   372
      Left            =   4920
      MaskColor       =   &H00000000&
      Picture         =   "RoutRTf05a.frx":0ED6
      Style           =   1  'Graphical
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3624
      Width           =   400
   End
   Begin VB.Frame z2 
      Height          =   30
      Left            =   180
      TabIndex        =   9
      Top             =   1440
      Width           =   5400
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "RoutRTf05a.frx":13C8
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdVew 
      DownPicture     =   "RoutRTf05a.frx":1B76
      Height          =   320
      Left            =   4560
      Picture         =   "RoutRTf05a.frx":2050
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "Parts Assigned To This Routing"
      Top             =   640
      Width           =   360
   End
   Begin VB.CommandButton cmdRen 
      Caption         =   "&Auto No"
      Enabled         =   0   'False
      Height          =   315
      Left            =   4800
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "Renumber In Current Order"
      Top             =   1800
      Width           =   825
   End
   Begin VB.ComboBox cmbRte 
      Height          =   315
      Left            =   1170
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   630
      Width           =   3345
   End
   Begin VB.ListBox lstOps 
      Height          =   2205
      Left            =   1170
      TabIndex        =   1
      Top             =   1620
      Width           =   3435
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   4800
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   0
      Top             =   4080
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4260
      FormDesignWidth =   5670
   End
   Begin VB.Image Dsup 
      Height          =   300
      Left            =   360
      Picture         =   "RoutRTf05a.frx":29C2
      Top             =   3600
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Enup 
      Height          =   300
      Left            =   840
      Picture         =   "RoutRTf05a.frx":2EB4
      Top             =   3600
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Endn 
      Height          =   300
      Left            =   600
      Picture         =   "RoutRTf05a.frx":33A6
      Top             =   3600
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Dsdn 
      Height          =   300
      Left            =   120
      Picture         =   "RoutRTf05a.frx":3898
      Top             =   3600
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Routing"
      Height          =   285
      Index           =   0
      Left            =   180
      TabIndex        =   5
      Top             =   630
      Width           =   915
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   285
      Index           =   4
      Left            =   180
      TabIndex        =   4
      Top             =   960
      Width           =   1005
   End
   Begin VB.Label txtDsc 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1170
      TabIndex        =   2
      Top             =   960
      Width           =   3075
   End
End
Attribute VB_Name = "RoutRTf05a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'12/29/05 Fixed null value in GetOperations
Option Explicit
'Dim RdoStm As rdoQuery
Dim AdoCmdStm As ADODB.Command

Dim bGoodRout As Byte
Dim bOnLoad As Byte

Dim iTotalOps As Integer
Dim iIndex As Integer

Dim sCurrRout As String

Dim sOpnum(300) As String * 4
Dim sOpcmt(300) As String * 30

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub cmbRte_Click()
   bGoodRout = GetRout(False)
   GetOperations (False)
   
End Sub


Private Sub cmbRte_LostFocus()
   cmbRte = CheckLen(cmbRte, 30)
   If Len(cmbRte) = 0 Then
      On Error Resume Next
      cmdCan.SetFocus
      Exit Sub
   Else
      bGoodRout = GetRout(False)
      If bGoodRout Then
         GetOperations (True)
      Else
         cmdRen.Enabled = False
      End If
      
   End If
   
End Sub

Private Sub cmdCan_Click()
   Unload Me
   
End Sub


Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   cmdCan_Click
   
End Sub


Private Sub cmdDn_Click()
   Dim iList As Integer
   Dim sComment As String
   Dim sOperation As String
   
   cmdRen.Enabled = True
   sOperation = sOpnum(iIndex + 1)
   sComment = sOpcmt(iIndex + 1)
   
   sOpnum(iIndex + 1) = sOpnum(iIndex)
   sOpcmt(iIndex + 1) = sOpcmt(iIndex)
   sOpnum(iIndex) = sOperation
   sOpcmt(iIndex) = sComment
   
   lstOps.Clear
   For iList = 0 To iTotalOps - 1
      lstOps.AddItem sOpnum(iList) & sOpcmt(iList)
   Next
   lstOps.Selected(iIndex + 1) = True
   lstOps.SetFocus
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 3154
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub cmdRen_Click()
   Dim bResponse As Byte
   bResponse = MsgBox("Auto Number And Save Operations?", ES_YESQUESTION, Caption)
   If bResponse = vbNo Then
      On Error Resume Next
      lstOps.SetFocus
      Width = Width + 10
      Exit Sub
   Else
      cmdRen.Enabled = False
      AutoNumber
   End If
   
End Sub

Private Sub cmdUp_Click()
   Dim iList As Integer
   Dim sComment As String
   Dim sOperation As String
   cmdRen.Enabled = True
   
   sOperation = sOpnum(iIndex - 1)
   sComment = sOpcmt(iIndex - 1)
   
   sOpnum(iIndex - 1) = sOpnum(iIndex)
   sOpcmt(iIndex - 1) = sOpcmt(iIndex)
   sOpnum(iIndex) = sOperation
   sOpcmt(iIndex) = sComment
   
   lstOps.Clear
   For iList = 0 To iTotalOps - 1
      lstOps.AddItem sOpnum(iList) & sOpcmt(iList)
   Next
   lstOps.Selected(iIndex - 1) = True
   lstOps.SetFocus
   
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
      FillRoutings
      If Len(sCurrRout) Then cmbRte = sCurrRout
      bGoodRout = GetRout(False)
      GetOperations (False)
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   
   On Error Resume Next
   sSql = "SELECT OPREF,OPNO,OPCOMT FROM RtopTable WHERE OPREF= ?"
   
   Set AdoCmdStm = New ADODB.Command
   AdoCmdStm.CommandText = sSql
   
   Dim prmOPRef As ADODB.Parameter
   Set prmOPRef = New ADODB.Parameter
   prmOPRef.Type = adChar
   prmOPRef.Size = 30
   AdoCmdStm.Parameters.Append prmOPRef
   
   'Set RdoStm = clsADOCon.CreateQuery("", sSql)
   GetRoutingIncrementDefault
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Set AdoCmdStm = Nothing
   FormUnload
   Set RoutRTf05a = Nothing
   
End Sub



Private Function GetRout(bGetOps As Byte) As Byte
   Dim RdoRte As ADODB.Recordset
   Dim sRout As String
   lstOps.Clear
   sRout = Compress(cmbRte)
   GetRout = False
   cmdRen.Enabled = False
   cmdUp.Picture = Dsup
   cmdUp.Enabled = False
   cmdDn.Picture = Dsdn
   cmdDn.Enabled = False
   
   On Error GoTo DiaErr1
   sSql = "Qry_GetToolRout '" & sRout & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRte)
   If bSqlRows Then
      With RdoRte
         GetRout = True
         cmbRte = "" & Trim(!RTNUM)
         txtDsc = "" & Trim(!RTDESC)
         sCurrRout = sRout
         ClearResultSet RdoRte
      End With
      If bGetOps Then GetOperations (False)
   Else
      cmbRte = ""
      txtDsc = ""
      sCurrRout = ""
      GetRout = False
   End If
   Set RdoRte = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getrout"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Function GetOperations(bOpen As Byte)
   Dim RdoOps As ADODB.Recordset
   Dim iList As Integer
   
   Erase sOpnum
   Erase sOpcmt
   iTotalOps = 0
   iIndex = 0
   iList = -1
   
   On Error GoTo DiaErr1
   
   'If bOpen Then RdoStm.MaxRows = 0 Else RdoStm.MaxRows = 15
   AdoCmdStm.Parameters(0).value = Compress(cmbRte)
'   RdoStm(0) = Compress(cmbRte)
   bSqlRows = clsADOCon.GetQuerySet(RdoOps, AdoCmdStm, ES_STATIC)
   If bSqlRows Then
      If bOpen Then RdoOps.MaxRecords = 0 Else RdoOps.MaxRecords = 15
      With RdoOps
         Do Until .EOF
            iList = iList + 1
            iTotalOps = iTotalOps + 1
            sOpnum(iList) = Format(!OPNO, "000")
            sOpcmt(iList) = "" & Trim(!OPCOMT)
            .MoveNext
         Loop
         ClearResultSet RdoOps
      End With
   End If
   For iList = 0 To iTotalOps
      sOpcmt(iList) = TrimComment(sOpcmt(iList))
      lstOps.AddItem sOpnum(iList) + sOpcmt(iList)
   Next
   If bOpen Then
      iIndex = 0
      cmdUp.Picture = Dsup
      cmdUp.Enabled = False
      If iTotalOps > 1 Then
         cmdDn.Picture = Endn
         cmdDn.Enabled = True
      End If
   End If
   Set RdoOps = Nothing
   On Error Resume Next
   If lstOps.ListCount > 0 And bOpen Then
      cmdRen.Enabled = True
      lstOps.Selected(0) = True
      lstOps.SetFocus
   Else
      cmdRen.Enabled = False
   End If
   Exit Function
   
DiaErr1:
   sProcName = "getopera"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub lstOps_Click()
   If lstOps.ListIndex < 0 Then lstOps.ListIndex = 0
   iIndex = lstOps.ListIndex
   If iTotalOps > 1 Then
      If iIndex > 0 Then
         cmdUp.Picture = Enup
         cmdUp.Enabled = True
      Else
         cmdUp.Picture = Dsup
         cmdUp.Enabled = False
      End If
      If iIndex + 1 >= iTotalOps Then
         cmdDn.Enabled = False
         cmdDn.Picture = Dsdn
      Else
         cmdDn.Enabled = True
         cmdDn.Picture = Endn
      End If
   End If
End Sub


Private Function TrimComment(sComment As String)
   Dim n As Integer
   On Error GoTo DiaErr1
   If Len(sComment) > 0 Then
      If Len(sComment) > 20 Then sComment = Left(sComment, 20)
      If Asc(Left(sComment, 1)) = 13 Then
         sComment = Right(sComment, Len(sComment) - 2)
      End If
      If Len(sComment) > 2 Then
         n = InStr(3, sComment, Chr(13))
         If n > 0 Then
            TrimComment = Left$(sComment, n - 1)
         Else
            TrimComment = sComment
         End If
      Else
         TrimComment = sComment
      End If
      If Len(TrimComment) < 20 Then
         TrimComment = TrimComment & (String$(20 - (Len(TrimComment)), " "))
      End If
   Else
      TrimComment = ""
   End If
   Exit Function
   
DiaErr1:
   sProcName = "trimcomme"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub AutoNumber()
   Dim iList As Integer
   Dim iNewNumber As Integer
   Dim iNewOps(300, 3) As Integer
   
   On Error GoTo DiaErr1
   cmdCan.Enabled = False
   MouseCursor 11
   For iList = 0 To lstOps.ListCount - 1
      iNewOps(iList, 0) = (iList + 1052)
      iNewOps(iList, 1) = Val(Left(lstOps.List(iList), 3))
   Next
   clsADOCon.BeginTrans
   For iList = 0 To lstOps.ListCount - 1
      sSql = "UPDATE RtopTable SET OPNO=" & str(iNewOps(iList, 0)) & " WHERE OPREF='" & sCurrRout & "' AND OPNO=" & str(iNewOps(iList, 1))
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
   Next
   clsADOCon.CommitTrans
   
   clsADOCon.BeginTrans
   iNewNumber = 0
   For iList = 0 To lstOps.ListCount - 1
      iNewNumber = iNewNumber + iAutoIncr
      sSql = "UPDATE RtopTable SET OPNO=" & str(iNewNumber) & " WHERE OPREF='" & sCurrRout & "' AND OPNO=" & str(iNewOps(iList, 0))
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
   Next
   clsADOCon.CommitTrans
   MouseCursor 0
   SysMsg "Auto Numbering Complete.", True, Me
   bGoodRout = GetRout(True)
   On Error Resume Next
   cmdCan.Enabled = True
   Exit Sub
   
DiaErr1:
   sProcName = "autonumber"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   On Error Resume Next
   MouseCursor 0
   clsADOCon.RollbackTrans
   cmdCan.Enabled = True
   DoModuleErrors Me
   
End Sub
