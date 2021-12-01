VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form RoutRTf04a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Merge Routings"
   ClientHeight    =   4665
   ClientLeft      =   2355
   ClientTop       =   1455
   ClientWidth     =   6465
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4665
   ScaleWidth      =   6465
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   1800
      TabIndex        =   24
      Top             =   240
      Width           =   3135
      Begin VB.OptionButton optNewRt 
         Caption         =   "Create New Routing"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Value           =   -1  'True
         Width           =   2775
      End
      Begin VB.OptionButton optAppRt 
         Caption         =   "Append/Insert to existing Routing"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   600
         Width           =   2895
      End
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "RoutRTf04a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   22
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.TextBox txtAtp 
      Height          =   285
      Left            =   1800
      TabIndex        =   8
      Tag             =   "1"
      Top             =   3690
      Width           =   555
   End
   Begin VB.TextBox txtRte 
      Height          =   285
      Left            =   1800
      TabIndex        =   0
      Tag             =   "3"
      Top             =   1350
      Width           =   3075
   End
   Begin VB.TextBox txtN2e 
      Height          =   285
      Left            =   3600
      TabIndex        =   7
      Tag             =   "1"
      Top             =   3330
      Width           =   555
   End
   Begin VB.TextBox txtN2b 
      Height          =   285
      Left            =   1800
      TabIndex        =   6
      Tag             =   "1"
      Top             =   3330
      Width           =   555
   End
   Begin VB.TextBox txtN1e 
      Height          =   285
      Left            =   3600
      TabIndex        =   4
      Top             =   2520
      Width           =   555
   End
   Begin VB.TextBox txtN1b 
      Height          =   285
      Left            =   1800
      TabIndex        =   3
      Top             =   2520
      Width           =   555
   End
   Begin VB.ComboBox cmbNw2 
      Height          =   315
      Left            =   1800
      Sorted          =   -1  'True
      TabIndex        =   5
      Tag             =   "3"
      Top             =   2970
      Width           =   3345
   End
   Begin VB.ComboBox cmbNw1 
      Height          =   315
      Left            =   1800
      Sorted          =   -1  'True
      TabIndex        =   2
      Tag             =   "3"
      Top             =   2160
      Width           =   3345
   End
   Begin VB.TextBox txtDsc 
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Tag             =   "2"
      Text            =   " "
      Top             =   1710
      Width           =   3075
   End
   Begin VB.CommandButton cmdMrg 
      Caption         =   "&Merge"
      Enabled         =   0   'False
      Height          =   315
      Left            =   5490
      TabIndex        =   10
      ToolTipText     =   "Merge The Routings"
      Top             =   2160
      Width           =   875
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5490
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5880
      Top             =   3960
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4665
      FormDesignWidth =   6465
   End
   Begin MSComctlLib.ProgressBar prg1 
      Height          =   300
      Left            =   1800
      TabIndex        =   23
      Top             =   4080
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "At Operation"
      Height          =   285
      Index           =   10
      Left            =   180
      TabIndex        =   21
      Top             =   3690
      Width           =   1275
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Entry(s) Blank Or 0 For All"
      Height          =   285
      Index           =   9
      Left            =   4230
      TabIndex        =   20
      Top             =   3330
      Width           =   2085
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Entry(s) Blank Or 0 For All"
      Height          =   285
      Index           =   8
      Left            =   4230
      TabIndex        =   19
      Top             =   2520
      Width           =   2085
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "To Operation"
      Height          =   285
      Index           =   7
      Left            =   2520
      TabIndex        =   18
      Top             =   3330
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "From Operation"
      Height          =   285
      Index           =   6
      Left            =   180
      TabIndex        =   17
      Top             =   3360
      Width           =   1515
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "To Operation"
      Height          =   285
      Index           =   5
      Left            =   2520
      TabIndex        =   16
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "From Operation"
      Height          =   285
      Index           =   4
      Left            =   180
      TabIndex        =   15
      Top             =   2520
      Width           =   1275
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "With Routing"
      Height          =   285
      Index           =   3
      Left            =   180
      TabIndex        =   14
      Top             =   2970
      Width           =   1635
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Merge Routing"
      Height          =   285
      Index           =   2
      Left            =   180
      TabIndex        =   13
      Top             =   2160
      Width           =   1635
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   285
      Index           =   1
      Left            =   180
      TabIndex        =   12
      Top             =   1710
      Width           =   1635
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "New Routing Number"
      Height          =   285
      Index           =   0
      Left            =   180
      TabIndex        =   11
      Top             =   1350
      Width           =   1635
   End
End
Attribute VB_Name = "RoutRTf04a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'12/2/05 Modified Merge
Option Explicit
Dim bOnLoad As Byte
Dim bShow As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub cmbNw1_LostFocus()
   cmbNw1 = CheckLen(cmbNw1, 30)
   
End Sub

Private Sub cmbNw2_LostFocus()
   cmbNw2 = CheckLen(cmbNw2, 30)
   
End Sub

Private Sub cmdCan_Click()
   Unload Me
   
End Sub


Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 3153
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub cmdMrg_Click()
   
   If (optAppRt) Then
      AppendToExistRt
   ElseIf optNewRt Then
      MergeCreateNewRt
   Else
      MsgBox "Select a type of routing merge.", vbInformation, Caption
   End If

End Sub

Private Sub AppendToExistRt()
   Dim RdoRte As ADODB.Recordset
   Dim RdoMrg As ADODB.Recordset
   
   Dim bResponse As Byte
   Dim bGoodNew As Byte
   Dim bGoodRt1 As Byte
   Dim bgoodRt2 As Byte
   
   Dim a As Integer
   Dim iList As Integer
   Dim n As Integer
   Dim iAtop As Integer
   Dim iEndOp As Integer
   Dim icpyOptCnt As Integer
   Dim iOldOpNo As Integer
   
   Dim sNewRout As String
   Dim sOldOne As String
   Dim sOldTwo As String
   Dim sSqlRte As String
   
   Dim sComments(300) As String
   Dim iOpNumbers(300) As Integer
   
   If Val(txtAtp) = 0 Then txtAtp = Format(txtN2b, "##0")
   bGoodRt1 = GetRout(Compress(cmbNw1))
   If Not bGoodRt1 Then
      MsgBox "The First Routing Wasn't Found.", vbInformation, Caption
      On Error Resume Next
      cmbNw1.SetFocus
      Exit Sub
   End If
   bgoodRt2 = GetRout(Compress(cmbNw2))
   If Not bgoodRt2 Then
      MsgBox "The Second Routing Wasn't Found.", vbInformation, Caption
      On Error Resume Next
      cmbNw2.SetFocus
      Exit Sub
   End If
   bResponse = MsgBox("Are You Sure That You Want To Merge The Routings.", ES_YESQUESTION, Caption)
   If bResponse = vbNo Then
      On Error Resume Next
      cmdCan.SetFocus
      Width = Width + 10
      Exit Sub
   End If
   prg1.Visible = True
   MouseCursor 11
   cmdMrg.Enabled = False
   cmdCan.Enabled = False
   sOldOne = Compress(cmbNw1)
   sOldTwo = Compress(cmbNw2)
   
   iAtop = Val(txtAtp)
   
   On Error GoTo DiaErr1
   
   ' Find out how may ops to copy
   sSql = "SELECT Count(OPNO) as cntOpts FROM RtopTable WHERE OPREF='" & sOldOne & "'" & _
            " AND OPNO >=" & txtN1b & " AND OPNO <=" & str(txtN1e)
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoMrg, ES_STATIC)
   prg1.Value = 30
   If bSqlRows Then
      icpyOptCnt = RdoMrg!cntOpts
      ClearResultSet RdoMrg
   Else
      MsgBox "The Merge Routing Operations couldn't Found.", vbInformation, Caption
      On Error Resume Next
      cmbNw2.SetFocus
      Exit Sub
   End If
   Set RdoMrg = Nothing
   
   ' Move the old opts to new
   sSql = "SELECT OPNO FROM RtopTable WHERE OPREF='" & sOldTwo & "'" & _
            "AND OPNO > " & txtAtp & " ORDER BY OPNO"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoMrg, ES_STATIC)
   
   clsADOCon.ADOErrNum = 0
   clsADOCon.BeginTrans
   
   iList = txtAtp + (icpyOptCnt * iAutoIncr)
   prg1.Value = 30
   If bSqlRows Then
      Do Until RdoMrg.EOF
      
         iList = iList + iAutoIncr
         iOldOpNo = RdoMrg!OPNO
         sSql = "Update RtopTable SET OPNO = " & iList & " WHERE OPREF='" & sOldTwo & "'" & _
                   " AND OPNO = " & iOldOpNo
         
         clsADOCon.ExecuteSQL sSql
         
         RdoMrg.MoveNext
      Loop
      ClearResultSet RdoMrg
   End If
   Set RdoMrg = Nothing
      
   ' Create a recordset to a merge/append rows.
   sSqlRte = "SELECT * FROM RtopTable WHERE OPREF='" & sOldTwo & "'"
   'Set RdoRte = RdoCon.OpenResultset(sSqlRte, rdOpenKeyset, rdConcurRowVer)
   Set RdoRte = clsADOCon.GetRecordSet(sSqlRte, adOpenKeyset) ' Not using Lock Type
      
   'Merge Routing
   ' reset the Opt number
   iList = txtAtp
   If Val(txtN1e) = 0 Then iEndOp = 1000 Else iEndOp = Val(txtN1e)
   sSql = "SELECT * FROM RtopTable WHERE OPREF='" & sOldOne & "' AND OPNO >= " & txtN1b & " AND OPNO <= " & str(iEndOp)
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoMrg, ES_STATIC)
   prg1.Value = 30
   If bSqlRows Then
      Do Until RdoMrg.EOF
         iList = iList + iAutoIncr
         n = n + 1
         sComments(n) = "" & Trim(RdoMrg!OPCOMT)
         iOpNumbers(n) = iList
         RdoRte.AddNew
         RdoRte!OPREF = "" & sOldTwo
         RdoRte!OPNO = iList
         RdoRte!OPSHOP = "" & Trim(RdoMrg!OPSHOP)
         RdoRte!OPCENTER = "" & Trim(RdoMrg!OPCENTER)
         RdoRte!OPSETUP = RdoMrg!OPSETUP
         RdoRte!OPUNIT = RdoMrg!OPUNIT
         RdoRte!OPQHRS = RdoMrg!OPQHRS
         RdoRte!OPMHRS = RdoMrg!OPMHRS
         RdoRte!OPSERVICE = RdoMrg!OPSERVICE
         RdoRte!OPSERVPART = "" & Trim(RdoMrg!OPSERVPART)
         RdoRte.Update
         RdoMrg.MoveNext
      Loop
      ClearResultSet RdoMrg
   End If
   
   prg1.Value = 80
   For iList = 1 To n
      sSql = "UPDATE RtopTable SET OPCOMT='" & sComments(iList) & "' WHERE OPREF='" & sOldTwo & "' AND OPNO=" & str(iOpNumbers(iList))
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
   Next
   On Error Resume Next
   Set RdoMrg = Nothing
   Set RdoRte = Nothing
   
   prg1.Value = 100
   MouseCursor 0
   
   If clsADOCon.ADOErrNum = 0 Then
      clsADOCon.CommitTrans
      SysMsg "Routing " & sOldOne & " Merged.", True, Me
      
      RoutRTe01a.cmbRte = sOldTwo
      RoutRTe01a.Show
      RoutRTe01a.txtDsc = ""
      
      cmdMrg.Enabled = True
      cmdCan.Enabled = True
      MouseCursor 13
      'sCurrRout = txtRte
      Unload Me
      
   Else
      clsADOCon.RollbackTrans
      clsADOCon.ADOErrNum = 0
      SysMsg "Routing " & sOldOne & " Couldn't get Merged.", True, Me
   End If
   
   cmdMrg.Enabled = True
   cmdCan.Enabled = True
   MouseCursor 13
   bShow = 1
   Exit Sub
   
DiaErr1:
   CurrError.Description = Err.Description
   Resume DiaErr2
DiaErr2:
   On Error Resume Next
   Set RdoRte = Nothing
   RdoMrg.Close
   cmdMrg.Enabled = True
   cmdCan.Enabled = True
   MouseCursor 0
   MsgBox CurrError.Description & " Couldn't Merge Routings.", vbExclamation, Caption

End Sub

Private Sub MergeCreateNewRt()
   Dim RdoRte As ADODB.Recordset
   Dim RdoMrg As ADODB.Recordset
   
   Dim bResponse As Byte
   Dim bGoodNew As Byte
   Dim bGoodRt1 As Byte
   Dim bgoodRt2 As Byte
   
   Dim a As Integer
   Dim iList As Integer
   Dim n As Integer
   Dim iAtop As Integer
   Dim iEndOp As Integer
   
   
   Dim sNewRout As String
   Dim sOldOne As String
   Dim sOldTwo As String
   Dim sSqlRte As String
   
   Dim sComments(300) As String
   Dim iOpNumbers(300) As Integer
   
   If Trim(txtRte) = "" Then
      MsgBox "Requires A Valid New Routing.", _
         vbInformation, Caption
      Exit Sub
   End If
   If Trim(txtDsc) = "" Then
      MsgBox "Requires A Valid New Routing Description.", _
         vbInformation, Caption
      Exit Sub
   End If
   
   bGoodNew = GetRout(Compress(txtRte))
   If bGoodNew Then
      MsgBox "That Routing Is Already Recorded.", vbInformation, Caption
      On Error Resume Next
      txtRte.SetFocus
      Exit Sub
   End If
   If Val(txtAtp) = 0 Then txtAtp = Format(txtN2b, "##0")
   bGoodRt1 = GetRout(Compress(cmbNw1))
   If Not bGoodRt1 Then
      MsgBox "The First Routing Wasn't Found.", vbInformation, Caption
      On Error Resume Next
      cmbNw1.SetFocus
      Exit Sub
   End If
   bgoodRt2 = GetRout(Compress(cmbNw2))
   If Not bgoodRt2 Then
      MsgBox "The Second Routing Wasn't Found.", vbInformation, Caption
      On Error Resume Next
      cmbNw2.SetFocus
      Exit Sub
   End If
   bResponse = MsgBox("Are You Sure That You Want To Merge The Routings.", ES_YESQUESTION, Caption)
   If bResponse = vbNo Then
      On Error Resume Next
      cmdCan.SetFocus
      Width = Width + 10
      Exit Sub
   End If
   prg1.Visible = True
   MouseCursor 11
   cmdMrg.Enabled = False
   cmdCan.Enabled = False
   sNewRout = Compress(txtRte)
   sOldOne = Compress(cmbNw1)
   sOldTwo = Compress(cmbNw2)
   
   iAtop = Val(txtAtp)
   
   On Error GoTo DiaErr1
   sSql = "INSERT RthdTable (RTREF,RTNUM,RTDESC) VALUES('" & sNewRout & "','" & txtRte & "','" & txtDsc & "')"
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   prg1.Value = 10
   sSqlRte = "SELECT * FROM RtopTable WHERE OPREF='" & sNewRout & "'"
   
   'Set RdoRte = RdoCon.OpenResultset(sSqlRte, rdOpenKeyset, rdConcurRowVer)
   Set RdoRte = clsADOCon.GetRecordSet(sSqlRte, adOpenKeyset) ' Not using Lock Type
   ' TODO Set RdoRte = RdoCon.OpenResultset(sSqlRte, rdOpenKeyset, rdConcurRowVer)
   
   'Merge Routing
   If Val(txtN1e) = 0 Then iEndOp = 1000 Else iEndOp = Val(txtN1e)
   sSql = "SELECT * FROM RtopTable WHERE OPREF='" & sOldOne & "' AND OPNO>=" & txtN1b & " AND OPNO<=" & str(iEndOp)
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoMrg, ES_STATIC)
   prg1.Value = 30
   If bSqlRows Then
      Do Until RdoMrg.EOF
         iList = iList + iAutoIncr
         n = n + 1
         sComments(n) = "" & Trim(RdoMrg!OPCOMT)
         iOpNumbers(n) = iList
         RdoRte.AddNew
         RdoRte!OPREF = "" & sNewRout
         RdoRte!OPNO = iList
         RdoRte!OPSHOP = "" & Trim(RdoMrg!OPSHOP)
         RdoRte!OPCENTER = "" & Trim(RdoMrg!OPCENTER)
         RdoRte!OPSETUP = RdoMrg!OPSETUP
         RdoRte!OPUNIT = RdoMrg!OPUNIT
         RdoRte!OPQHRS = RdoMrg!OPQHRS
         RdoRte!OPMHRS = RdoMrg!OPMHRS
         RdoRte!OPSERVICE = RdoMrg!OPSERVICE
         RdoRte!OPSERVPART = "" & Trim(RdoMrg!OPSERVPART)
         RdoRte.Update
         RdoMrg.MoveNext
      Loop
      ClearResultSet RdoMrg
   End If
   
   'With Routing First step
   a = 0
   If iAtop >= iEndOp Then
      'Append it
      a = 1
      If Val(txtN2e) = 0 Then iEndOp = 1000 Else iEndOp = Val(txtN2e)
      sSql = "SELECT * FROM RtopTable WHERE OPREF='" & sOldTwo & "' AND OPNO>=" & txtN1b & " AND OPNO<=" & str(iEndOp)
   Else
      'find where to merge
      sSql = "SELECT * FROM RtopTable WHERE OPREF='" & sOldTwo & "' AND OPNO>=" & txtN1b & " AND OPNO<=" & str(iAtop)
   End If
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoMrg, ES_STATIC)
   prg1.Value = 60
   If bSqlRows Then
      Do Until RdoMrg.EOF
         iList = iList + iAutoIncr
         n = n + 1
         sComments(n) = "" & Trim(RdoMrg!OPCOMT)
         iOpNumbers(n) = iList
         RdoRte.AddNew
         RdoRte!OPREF = "" & sNewRout
         RdoRte!OPNO = iList
         RdoRte!OPSHOP = "" & Trim(RdoMrg!OPSHOP)
         RdoRte!OPCENTER = "" & Trim(RdoMrg!OPCENTER)
         RdoRte!OPSETUP = RdoMrg!OPSETUP
         RdoRte!OPUNIT = RdoMrg!OPUNIT
         RdoRte!OPQHRS = RdoMrg!OPQHRS
         RdoRte!OPMHRS = RdoMrg!OPMHRS
         RdoRte!OPSERVICE = RdoMrg!OPSERVICE
         RdoRte!OPSERVPART = "" & Trim(RdoMrg!OPSERVPART)
         RdoRte.Update
         RdoMrg.MoveNext
      Loop
      ClearResultSet RdoMrg
   End If
   prg1.Value = 70
   If a = 0 Then
      'Wants to merge behind here
      sSql = "SELECT * FROM RtopTable WHERE OPREF='" & sOldTwo & "' AND OPNO>" & str(iAtop)
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoMrg, ES_STATIC)
      If bSqlRows Then
         Do Until RdoMrg.EOF
            iList = iList + iAutoIncr
            n = n + 1
            sComments(n) = "" & Trim(RdoMrg!OPCOMT)
            iOpNumbers(n) = iList
            RdoRte.AddNew
            RdoRte!OPREF = "" & sNewRout
            RdoRte!OPNO = iList
            RdoRte!OPSHOP = "" & Trim(RdoMrg!OPSHOP)
            RdoRte!OPCENTER = "" & Trim(RdoMrg!OPCENTER)
            RdoRte!OPSETUP = RdoMrg!OPSETUP
            RdoRte!OPUNIT = RdoMrg!OPUNIT
            RdoRte!OPQHRS = RdoMrg!OPQHRS
            RdoRte!OPMHRS = RdoMrg!OPMHRS
            RdoRte!OPSERVICE = RdoMrg!OPSERVICE
            RdoRte!OPSERVPART = "" & Trim(RdoMrg!OPSERVPART)
            RdoRte.Update
            RdoMrg.MoveNext
         Loop
         ClearResultSet RdoMrg
      End If
   End If
   prg1.Value = 80
   For iList = 1 To n
      sSql = "UPDATE RtopTable SET OPCOMT='" & sComments(iList) & "' WHERE OPREF='" & sNewRout & "' AND OPNO=" & str(iOpNumbers(iList))
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
   Next
   On Error Resume Next
   Set RdoMrg = Nothing
   Set RdoRte = Nothing
   prg1.Value = 100
   MouseCursor 0
   SysMsg "Routing " & txtRte & " Merged.", True, Me
   cmdMrg.Enabled = True
   cmdCan.Enabled = True
   MouseCursor 13
   bShow = 1
   sCurrRout = txtRte
   RoutRTe01a.cmbRte = txtRte
   RoutRTe01a.Show
   RoutRTe01a.txtDsc = txtDsc
   Unload Me
   Exit Sub
   
DiaErr1:
   CurrError.Description = Err.Description
   Resume DiaErr2
DiaErr2:
   On Error Resume Next
   Set RdoRte = Nothing
   RdoMrg.Close
   cmdMrg.Enabled = True
   cmdCan.Enabled = True
   clsADOCon.ExecuteSQL "DELETE FROM RthdTable WHERE RTREF='" & sNewRout & "'"
   clsADOCon.ExecuteSQL ("DELETE FROM RtopTable WHERE OPREF='" & sNewRout & "'")
   MouseCursor 0
   MsgBox CurrError.Description & " Couldn't Merge Routings.", vbExclamation, Caption
   
End Sub

Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      bOnLoad = 0
      FillCombos
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST
   FormatControls
   GetRoutingIncrementDefault
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   If bShow = 0 Then FormUnload
   Set RoutRTf04a = Nothing
   
End Sub


Private Sub optAppRt_Click()
   DisableNewRts (False)
End Sub

Private Sub optNewRt_Click()
   DisableNewRts (True)
End Sub

Private Function DisableNewRts(bDis As Boolean)

   If (bDis) Then
      txtRte.Enabled = bDis
      txtDsc.Enabled = bDis
      txtN2b.Enabled = bDis
      txtN2e.Enabled = bDis
      
   Else
      txtRte = ""
      txtDsc = ""
      txtN2b = ""
      txtN2e = ""
      txtRte.Enabled = bDis
      txtDsc.Enabled = bDis
      txtN2b.Enabled = bDis
      txtN2e.Enabled = bDis
      
      cmdMrg.Enabled = True
   
   End If
End Function

Private Sub txtAtp_LostFocus()

   If (optNewRt) Then
      txtAtp = CheckLen(txtAtp, 3)
      txtAtp = Format(Abs(Val(txtAtp)), "##0")
      If Val(txtAtp) < Val(txtN2b) Then txtAtp = txtN2b
      If Val(txtAtp) > Val(txtN2e) Then txtAtp = txtN2e
   End If
End Sub

Private Sub txtDsc_LostFocus()
   txtDsc = CheckLen(txtDsc, 30)
   txtDsc = StrCase(txtDsc)
   If Len(txtDsc) > 0 And Len(txtRte) > 0 Then
      cmdMrg.Enabled = True
   Else
      cmdMrg.Enabled = False
   End If
   
End Sub

Private Function GetRout(sRouting As String) As Byte
   Dim RdoRte As ADODB.Recordset
   Dim sRout As String
   
   sRout = Compress(sRout)
   GetRout = False
   On Error GoTo DiaErr1
   sSql = "SELECT RTREF FROM RthdTable WHERE RTREF='" & sRouting & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRte, ES_STATIC)
   If bSqlRows Then
      GetRout = True
   Else
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

Private Sub txtN1b_LostFocus()
   txtN1b = CheckLen(txtN1b, 3)
   txtN1b = Format(Abs(Val(txtN1b)), "##0")
   If Val(txtN1b) = 0 Then Exit Sub
   If Val(txtN1e) = 0 Then Exit Sub
   If Val(txtN1b) > Val(txtN1e) Then txtN1e = txtN1b
   
End Sub

Private Sub txtN1e_LostFocus()
   txtN1e = CheckLen(txtN1e, 3)
   txtN1e = Format(Abs(Val(txtN1e)), "##0")
   If Val(txtN1e) = 0 Then Exit Sub
   If Val(txtN1b) = 0 Then Exit Sub
   If Val(txtN1e) < Val(txtN1b) Then txtN1b = txtN1e
   
End Sub

Private Sub txtN2b_LostFocus()
   txtN2b = CheckLen(txtN2b, 3)
   txtN2b = Format(Abs(Val(txtN2b)), "##0")
   If Val(txtN2e) = 0 Then Exit Sub
   If Val(txtN2b) = 0 Then Exit Sub
   If Val(txtN2b) > Val(txtN2e) Then txtN2e = txtN2b
   txtAtp = txtN2b
   
End Sub

Private Sub txtN2e_LostFocus()
   txtN2e = CheckLen(txtN2e, 3)
   txtN2e = Format(Abs(Val(txtN2e)), "##0")
   If Val(txtN2e) = 0 Then Exit Sub
   If Val(txtN2b) = 0 Then Exit Sub
   If Val(txtN2e) < Val(txtN2b) Then txtN2b = txtN1e
   
End Sub

Private Sub txtRte_Click()
   txtN1b = "0"
   txtN1e = "0"
   txtN2b = "0"
   txtN2e = "0"
   cmdMrg.Enabled = False
   
End Sub

Private Sub FillCombos()
   Dim RdoCmb As ADODB.Recordset
   
   On Error GoTo DiaErr1
   sSql = "Qry_FillRoutings"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCmb)
   If bSqlRows Then
      With RdoCmb
         cmbNw1 = "" & Trim(!RTNUM)
         cmbNw2 = "" & Trim(!RTNUM)
         Do Until .EOF
            AddComboStr cmbNw1.hwnd, "" & Trim(!RTNUM)
            AddComboStr cmbNw2.hwnd, "" & Trim(!RTNUM)
            .MoveNext
         Loop
         ClearResultSet RdoCmb
      End With
   End If
   Set RdoCmb = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombos"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub
