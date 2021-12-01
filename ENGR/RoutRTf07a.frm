VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form RoutRTf07a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Copy A Routing From A Manufacturing Order"
   ClientHeight    =   3405
   ClientLeft      =   2430
   ClientTop       =   1515
   ClientWidth     =   6585
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3405
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox optDefault 
      Caption         =   "Becomes the Default Routing"
      Height          =   192
      Left            =   1320
      TabIndex        =   4
      ToolTipText     =   "Make This The Default Routing For The Part"
      Top             =   2640
      Value           =   1  'Checked
      Width           =   2412
   End
   Begin VB.TextBox txtNewDesc 
      Height          =   288
      Left            =   1320
      TabIndex        =   3
      Tag             =   "2"
      Top             =   2280
      Width           =   3095
   End
   Begin VB.ComboBox cmbRun 
      Height          =   288
      Left            =   5520
      Sorted          =   -1  'True
      TabIndex        =   1
      Tag             =   "1"
      ToolTipText     =   "All Runs From This Part Number"
      Top             =   1200
      Width           =   855
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "RoutRTf07a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "C&opy"
      Enabled         =   0   'False
      Height          =   315
      Left            =   5520
      TabIndex        =   5
      ToolTipText     =   "Copy The Existing MO Routing To The New Routing"
      Top             =   1920
      Width           =   875
   End
   Begin VB.TextBox txtNew 
      Height          =   285
      Left            =   1320
      TabIndex        =   2
      Tag             =   "3"
      Top             =   1980
      Width           =   3095
   End
   Begin VB.ComboBox cmbRte 
      Height          =   315
      Left            =   1320
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Contains Parts With MO's"
      Top             =   1200
      Width           =   3345
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5520
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   0
      Top             =   3120
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3405
      FormDesignWidth =   6585
   End
   Begin MSComctlLib.ProgressBar prg1 
      Height          =   300
      Left            =   1320
      TabIndex        =   11
      Top             =   3000
      Width           =   3012
      _ExtentX        =   5318
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Run"
      Height          =   252
      Index           =   8
      Left            =   4800
      TabIndex        =   12
      Top             =   1200
      Width           =   672
   End
   Begin VB.Label txtDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   1320
      TabIndex        =   9
      Top             =   1560
      Width           =   3072
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "New Number"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   8
      Top             =   1980
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   7
      Top             =   1170
      Width           =   1215
   End
End
Attribute VB_Name = "RoutRTf07a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'1/24/07 New
Option Explicit
Dim bGoodOld As Byte
Dim bGoodNew As Byte
Dim bGoodRun As Byte
Dim bOnLoad As Byte
Dim bShow As Byte

Dim sOldRout As String
Dim sNewRout As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd


Private Sub cmbRte_Click()
   cmbRte = GetCurrentPart(cmbRte, txtDsc)
   FillTheseRuns
   
End Sub

Private Sub cmbRte_LostFocus()
   cmbRte = CheckLen(cmbRte, 30)
   cmbRte = GetCurrentPart(cmbRte, txtDsc)
   FillTheseRuns
   
End Sub


Private Sub cmbRun_LostFocus()
   Dim iList As Integer
   If cmbRun.ListCount > 0 Then
      bGoodRun = 0
      For iList = 0 To cmbRun.ListCount - 1
         If Val(cmbRun.List(iList)) = Val(cmbRun) Then bGoodRun = 1
      Next
   Else
      bGoodRun = 0
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
      OpenHelpContext 3156
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub cmdNew_Click()
   Dim bResponse As Byte
   Dim sMsg As String
   If bGoodRun = 0 Then
      MsgBox "Please Select A Valid Manufacturing Order Run.", _
         vbInformation, Caption
      Exit Sub
   End If
   bGoodNew = GetRout()
   If bGoodNew = 1 Then
      CopyRouting
   Else
      sMsg = "There Is A Previously Recorded Routing With This Number." & vbCrLf _
             & "Do You Wish To Continue And Overwrite The Existing Routing?"
      bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
      If bResponse = vbYes Then CopyRouting Else CancelTrans
   End If
   
End Sub


Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      bOnLoad = 0
      FillOrders
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
   Set RoutRTf07a = Nothing
   
End Sub




Private Sub txtDsc_Change()
   If Left(txtDsc, 5) <> "*** P" Then
      bGoodOld = 1
      txtNewDesc = txtDsc
      txtNew = cmbRte
      cmdNew.Enabled = True
   Else
      bGoodOld = 0
      cmdNew.Enabled = False
      txtNewDesc = ""
      txtNew = ""
   End If
   
End Sub

Private Sub txtNew_LostFocus()
   txtNew = CheckLen(txtNew, 30)
   If Len(txtNew) = 0 Then Exit Sub
   sOldRout = Compress(cmbRte)
   sNewRout = Compress(txtNew)
   
End Sub



Private Function GetRout() As Byte
   Dim RdoRte As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "Qry_GetToolRout '" & Compress(txtNew) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRte, ES_FORWARD)
   If bSqlRows Then GetRout = 0 Else GetRout = 1
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
   
   Dim txtRoutRevNotes As String    'bbs added on 3/23/2016
   Dim txtRoutBy As String  'bbs added 3/28/2016
   Dim txtRoutDate As String 'bbs added 3/31/20116
   
   bResponse = MsgBox("Copy Routing " & Trim(cmbRte) & " To " _
               & Trim(txtNew) & ".", ES_YESQUESTION, Caption)
   If bResponse = vbNo Then
      CancelTrans
      cmdCan.SetFocus
      Width = Width + 10
      Exit Sub
   End If
   
   Call GetRevNotes(txtRoutRevNotes, txtRoutBy, txtRoutDate)    'bbs added on 3/23/2016
   
   sNewRout = Compress(txtNew)
   sOldRout = Compress(cmbRte)
   
   MouseCursor 13
   cmdCan.Enabled = False
   prg1.Visible = True
   prg1.Value = 5
   
   
   On Error Resume Next
   'BBS Changed the query below on 7/9/2010 for Ticket #15214 (was OPUNIT and OPSETUP)
   sSql = "SELECT OPREF,OPRUN,OPNO,OPSHOP,OPCENTER,OPSUHRS,OPUNITHRS,OPPICKOP," _
          & "OPSERVPART,OPQHRS,OPMHRS,OPSVCUNIT,OPTOOLLIST,OPCOMT FROM RnopTable " _
          & "WHERE (OPREF='" & Compress(cmbRte) & "' AND OPRUN=" & cmbRun & ")"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCpy, ES_STATIC)
   If bSqlRows Then
      
      Dim opser As Integer
      Err.Clear
      clsADOCon.ADOErrNum = 0
      clsADOCon.BeginTrans
      
      sSql = "DELETE FROM RtopTable WHERE OPREF='" & Compress(txtNew) & "'"
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
      prg1.Value = 10
      
      sSql = "DELETE FROM RthdTable WHERE RTREF='" & Compress(txtNew) & "'"
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
      
      'bbs added the RTREVNOTES to the below insert query on 3/23/2016
      sSql = "INSERT INTO RthdTable (RTREF,RTNUM,RTDESC,RTREVNOTES, RTBY, RTDATE) VALUES('" _
             & Compress(txtNew) & "','" _
             & txtNew & "','" _
             & txtNewDesc & "','" _
             & txtRoutRevNotes & "','" _
             & txtRoutBy & "','" _
             & txtRoutDate & "')"
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
      prg1.Value = 20
      
      sSql = "SELECT * FROM RtopTable where OPREF='" & Compress(txtNew) & "'"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoRte, ES_KEYSET)
      With RdoCpy
         Do Until .EOF
            If prg1.Value < 95 Then prg1.Value = prg1.Value + 5
            
            If (Trim(!OPSERVPART) = "") Then
               opser = 0
            Else
               opser = 1
            End If
               
            RdoRte.AddNew
            RdoRte!OPREF = "" & Compress(txtNew)    'BBS Changed on 7/26/2010 for Ticket #15214 (was Trim(!OPREF))
            RdoRte!OPNO = !OPNO
            RdoRte!OPSHOP = "" & Trim(!OPSHOP)
            RdoRte!OPCENTER = "" & Trim(!OPCENTER)
            RdoRte!OPSETUP = !OPSUHRS   'BBS Changed on 7/9/2010 for Ticket #15214 (was OPSETUP)
            RdoRte!OPUNIT = !OPUNITHRS  'BBS Changed on 7/9/2010 for Ticket #15214 (was OPUNIT)
            RdoRte!OPPICKOP = !OPPICKOP
            RdoRte!OPSERVPART = "" & Trim(!OPSERVPART)
            RdoRte!OPSERVICE = CStr(opser)
            RdoRte!OPQHRS = !OPQHRS
            RdoRte!OPMHRS = !OPMHRS
            RdoRte!OPSVCUNIT = !OPSVCUNIT
            RdoRte!OPTOOLLIST = "" & Trim(!OPTOOLLIST)
            RdoRte!OPCOMT = "" & Trim(!OPCOMT)
            RdoRte.Update
            .MoveNext
         Loop
      End With
      prg1.Value = 100
      MouseCursor 0
      If clsADOCon.ADOErrNum = 0 Then
         clsADOCon.CommitTrans
         MsgBox "The Routing Was Successfully Copied.", _
            vbInformation, Caption
         If optDefault.Value = vbChecked Then
            clsADOCon.ExecuteSQL ("UPDATE PartTable SET PAROUTING='" _
                            & Compress(txtNew) & "' WHERE PARTNUM='" & Compress(cmbRte) & "'")
                            
         End If
      Else
         clsADOCon.RollbackTrans
         clsADOCon.ADOErrNum = 0
         MsgBox "The Routing Could Not Be Copied.", _
            vbInformation, Caption
      End If
   Else
      prg1.Visible = False
      MouseCursor 0
      MsgBox "No Matching Rows Were Found To Copy.", _
         vbInformation, Caption
   End If
   cmdCan.Enabled = True
   Set RdoCpy = Nothing
   Set RdoRte = Nothing
   prg1.Visible = False
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

Private Sub FillOrders()
   Dim RdoRns As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT DISTINCT PARTREF,PARTNUM,RUNREF FROM " _
          & "PartTable,RunsTable WHERE PARTREF=RUNREF ORDER BY PARTREF"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRns, ES_FORWARD)
   If bSqlRows Then
      With RdoRns
         Do Until .EOF
            AddComboStr cmbRte.hwnd, "" & Trim(!PartNum)
            .MoveNext
         Loop
         ClearResultSet RdoRns
      End With
      If cmbRte.ListCount > 0 Then
         cmbRte = cmbRte.List(0)
         cmbRte = GetCurrentPart(cmbRte, txtDsc)
         FillTheseRuns
      End If
   End If
   On Error Resume Next
   Set RdoRns = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "fillorders"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MDISect.ActiveForm
   
End Sub

Private Sub FillTheseRuns()
   Dim RdoMoRuns As ADODB.Recordset
   cmbRun.Clear
   bGoodRun = 0
   If Left(txtDsc = "*** P", 5) Then Exit Sub
   On Error GoTo DiaErr1
   sSql = "SELECT DISTINCT RUNNO FROM RunsTable WHERE RUNREF='" _
          & Compress(cmbRte) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoMoRuns, ES_FORWARD)
   If bSqlRows Then
      With RdoMoRuns
         Do Until .EOF
            AddComboStr cmbRun.hwnd, "" & Trim(.Fields(0))
            .MoveNext
         Loop
      End With
      ClearResultSet RdoMoRuns
   End If
   If cmbRun.ListCount > 0 Then
      cmbRun = cmbRun.List(0)
      If GetPreferenceValue("AutoSelectLastRun") = "1" Then cmbRun = cmbRun.List(cmbRun.ListCount - 1)
      bGoodRun = 1
   End If
   Set RdoMoRuns = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "filltheseruns"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MDISect.ActiveForm
   
End Sub

Private Sub txtNewDesc_LostFocus()
   txtNewDesc = CheckLen(txtNewDesc, 30)
   
End Sub


'bbs added this new function on 3/23/2016
Private Sub GetRevNotes(ByRef txtRevNotes As String, ByRef txtRoutBy As String, txtRoutDate As String)
   Dim RdoRevNotes As ADODB.Recordset
   On Error GoTo RevNoteErr
   sSql = "SELECT RTREVNOTES, RTBY, RTDATE FROM RthdTable WHERE RTREF = '" & Compress(cmbRte) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRevNotes, ES_FORWARD)
   If bSqlRows Then
        txtRevNotes = Trim(RdoRevNotes!RTREVNOTES) & ""
        txtRoutBy = Trim(RdoRevNotes!RTBY) & ""
        txtRoutDate = Trim(RdoRevNotes!RTDATE) & ""
   Else
        txtRevNotes = ""
        txtRoutBy = ""
        txtRoutDate = ""
   End If
    
   Set RdoRevNotes = Nothing
   Exit Sub
   
RevNoteErr:
   sProcName = "GetRevNotes"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub
