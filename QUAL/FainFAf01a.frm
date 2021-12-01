VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FainFAf01a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Copy A First Article Report"
   ClientHeight    =   3045
   ClientLeft      =   2685
   ClientTop       =   1425
   ClientWidth     =   7200
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   Icon            =   "FainFAf01a.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3045
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtNewDesc 
      Height          =   288
      Left            =   1440
      TabIndex        =   4
      Tag             =   "2"
      ToolTipText     =   "Report Description"
      Top             =   2280
      Width           =   3012
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "FainFAf01a.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   12
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdCpy 
      Caption         =   "C&opy"
      Height          =   315
      Left            =   6240
      TabIndex        =   11
      ToolTipText     =   "Copy The Old Report To The New Report"
      Top             =   600
      Width           =   855
   End
   Begin VB.TextBox txtRev 
      Height          =   285
      Left            =   5280
      TabIndex        =   3
      Tag             =   "3"
      Top             =   1920
      Width           =   735
   End
   Begin VB.TextBox txtNew 
      Height          =   285
      Left            =   1440
      TabIndex        =   2
      Tag             =   "3"
      Top             =   1920
      Width           =   3015
   End
   Begin VB.ComboBox cmbPrt 
      DataSource      =   "rDt1"
      Height          =   288
      Left            =   1440
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Enter a New Part or Select From List (30 chars)"
      Top             =   1080
      Width           =   3255
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6480
      Top             =   2520
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3045
      FormDesignWidth =   7200
   End
   Begin VB.ComboBox cmbRev 
      ForeColor       =   &H00800000&
      Height          =   288
      Left            =   5280
      Sorted          =   -1  'True
      TabIndex        =   1
      Tag             =   "3"
      Text            =   " "
      ToolTipText     =   "SelectRevision From List"
      Top             =   1080
      Width           =   945
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Index           =   0
      Left            =   6240
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin MSComctlLib.ProgressBar prg1 
      Height          =   300
      Left            =   1440
      TabIndex        =   13
      Top             =   2640
      Width           =   3972
      _ExtentX        =   7011
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Rev"
      Height          =   288
      Index           =   3
      Left            =   4800
      TabIndex        =   10
      Top             =   1920
      Width           =   708
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "New Report"
      Height          =   288
      Index           =   2
      Left            =   240
      TabIndex        =   9
      Top             =   1920
      Width           =   1548
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   1440
      TabIndex        =   8
      Top             =   1440
      Width           =   3012
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Report Number"
      Height          =   288
      Index           =   1
      Left            =   240
      TabIndex        =   7
      Top             =   1080
      Width           =   1548
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Rev"
      Height          =   288
      Index           =   0
      Left            =   4800
      TabIndex        =   6
      Top             =   1080
      Width           =   708
   End
End
Attribute VB_Name = "FainFAf01a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'New 5/16/03
'10/4/06 Completely revamped CopyReport
Option Explicit
Dim bGoodReport As Byte
Dim bCancel As Byte
Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub



Private Sub cmbPrt_Click()
   GetReport
   
End Sub

Private Sub cmbPrt_LostFocus()
   cmbPrt = CheckLen(cmbPrt, 30)
   If bCancel = 1 Then Exit Sub
   GetReport
   
End Sub


Private Sub cmbRev_Click()
   GetDescription
   
End Sub

Private Sub cmbRev_LostFocus()
   Dim b As Byte
   Dim iList As Integer
   
   If bCancel = 1 Or Trim(cmbRev) = "" Then Exit Sub
   For iList = 1 To cmbRev.ListCount - 1
      If cmbRev = cmbRev.List(1) Then b = 1
   Next
   If b = 0 Then
      Beep
      cmbRev = cmbRev.List(0)
   End If
   GetDescription
   
End Sub


Private Sub cmdCan_Click(Index As Integer)
   Unload Me
   
End Sub


Private Sub cmdCan_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   bCancel = 1
   
End Sub


Private Sub cmdCpy_Click()
   Dim b As Byte
   Dim bResponse As Byte
   Dim sMsg As String
   
   If Trim(txtNew) = "" Then
      MsgBox "Requires A Valid Report Number.", _
         vbInformation, Caption
      txtNew.SetFocus
      Exit Sub
   End If
   If Len(Trim(txtNew)) < 4 Then
      MsgBox "Requires A Report Number With 4 Or More Chars.", _
         vbInformation, Caption
      txtNew.SetFocus
      Exit Sub
   End If
   b = IllegalCharacters(txtNew)
   If b > 0 Then
      MsgBox "The Report Number Contains An Illegal " & Chr$(b) & ".", _
         vbExclamation, Caption
      Exit Sub
   End If
   
   'Okay, Still want to?
   b = CheckReport
   If b = 1 Then
      MsgBox "That Report Number And Revision Has Been Recorded.", _
         vbInformation, Caption
      On Error Resume Next
      txtNew.SetFocus
      Exit Sub
   End If
   sMsg = "Do You Want To Create A New First Article To" & vbCr _
          & "A New Report (Acceptance Won't be Copied)?"
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      CopyReport
   Else
      CancelTrans
   End If
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 6250
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub Form_Activate()
   Static b As Byte
   MdiSect.lblBotPanel = Caption
   If bOnLoad = 1 Then
      bOnLoad = 0
      b = FillCombo()
   End If
   MouseCursor 0
   If b = 0 Then
      MsgBox "There Are No First Article Reports Recorded.", _
         vbInformation, Caption
      Unload Me
   End If
   
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
   Set FainFAf01a = Nothing
   
End Sub

Private Function FillCombo() As Byte
   On Error GoTo DiaErr1
   cmbPrt.Clear
   sSql = "Qry_FillFirstArticles"
   LoadComboBox cmbPrt
   If cmbPrt.ListCount > 0 Then
      FillCombo = 1
      cmbPrt = cmbPrt.List(0)
      'bGoodReport = GetReport()
   Else
      FillCombo = 0
   End If
   Exit Function
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function


Private Function GetReport() As Byte
   Dim RdoRep As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT FA_REF,FA_NUMBER,FA_DESCRIPTION FROM " _
          & "FahdTable WHERE FA_REF='" & Compress(cmbPrt) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRep, ES_FORWARD)
   If bSqlRows Then
      With RdoRep
         cmbPrt = "" & Trim(!FA_NUMBER)
         lblDsc = "" & Trim(!FA_DESCRIPTION)
         txtNewDesc = lblDsc
         GetReport = 1
         ClearResultSet RdoRep
      End With
      FillRevisions
   Else
      GetReport = 0
      lblDsc = "*** Report Wasn't Found ***"
      txtNewDesc = ""
   End If
   Set RdoRep = Nothing
   Exit Function
   
DiaErr1:
   bGoodReport = 0
   sProcName = "getreport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub lblDsc_Change()
   If Left(lblDsc.Caption, 10) = "*** Report" Then
      lblDsc.ForeColor = ES_RED
   Else
      lblDsc.ForeColor = Es_TextForeColor
   End If
End Sub





Private Sub FillRevisions()
   cmbRev.Clear
   On Error GoTo DiaErr1
   sSql = "SELECT FA_REF,FA_REVISION FROM " _
          & "FahdTable WHERE FA_REF='" & Compress(cmbPrt) & "'"
   LoadComboBox cmbRev
   Exit Sub
   
DiaErr1:
   sProcName = "fillrevs"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Function CheckReport() As Byte
   Dim RdoChk As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT FA_REF,FA_REVISION FROM FahdTable WHERE " _
          & "(FA_REF='" & Compress(txtNew) & "' AND FA_REVISION='" _
          & Trim(txtRev) & "')"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoChk, ES_FORWARD)
   If bSqlRows Then CheckReport = 1
   ClearResultSet RdoChk
   Set RdoChk = Nothing
   Exit Function
   
DiaErr1:
   CheckReport = 0
   sProcName = "checkreport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub GetDescription()
   Dim RdoRep As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT FA_REF,FA_DESCRIPTION FROM " _
          & "FahdTable WHERE (FA_REF='" & Compress(cmbPrt) & "' " _
          & "AND FA_REVISION='" & cmbRev & "')"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRep, ES_FORWARD)
   If bSqlRows Then
      With RdoRep
         lblDsc = "" & Trim(!FA_DESCRIPTION)
         ClearResultSet RdoRep
      End With
   Else
      lblDsc = "*** Report Wasn't Found ***"
   End If
   Set RdoRep = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getdescript"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub CopyReport()
   Dim RdoCpy As ADODB.Recordset
   
   Dim sNewRpt As String
   Dim sNewRev As String
   Dim sOldRpt As String
   Dim sOldRev As String
   
   Dim vDate As Variant
   
   sNewRpt = Compress(txtNew)
   sNewRev = Trim(txtRev)
   sOldRpt = Compress(cmbPrt)
   sOldRev = Trim(cmbRev)
   vDate = Format(ES_SYSDATE, "mm/dd/yy")
   
   MouseCursor 13
   cmdCpy.Enabled = False
   
   'Header (In case the temp tables remain)
   On Error Resume Next
   prg1.Visible = True
   prg1.Value = 10
   sSql = "DROP TABLE ##Fahd"
   clsADOCon.ExecuteSql sSql
   
   sSql = "DROP TABLE ##Fait"
   clsADOCon.ExecuteSql sSql
   
   sSql = "DROP TABLE ##Fadc"
   clsADOCon.ExecuteSql sSql
   
   Err.Clear
   clsADOCon.ADOErrNum = 0
   clsADOCon.BeginTrans
   'FahdTable
   sSql = "SELECT * INTO ##Fahd from FahdTable where FA_REF='" & sOldRpt & "' " _
          & "AND FA_REVISION='" & sOldRev & "'"
   clsADOCon.ExecuteSql sSql
   
   sSql = "UPDATE ##Fahd SET FA_REF='" & sNewRpt & "',FA_NUMBER='" & txtNew _
          & "',FA_REVISION='" & sNewRev & "',FA_DESCRIPTION='" & txtNewDesc _
          & "',FA_MORUNPART='',FA_MORUNNO='',FA_CREATED='" & vDate & "',FA_REVISED='" _
          & vDate & "',FA_INSPECTED=NULL,FA_PRINTED=0"
   clsADOCon.ExecuteSql sSql
   
   sSql = "INSERT INTO FahdTable SELECT * FROM ##Fahd"
   clsADOCon.ExecuteSql sSql
   prg1.Value = 30
   Sleep 200
   
   'FaitTable
   sSql = "SELECT * INTO ##Fait from FaitTable where FA_ITNUMBER='" & sOldRpt & "' " _
          & "AND FA_ITREVISION='" & sOldRev & "'"
   clsADOCon.ExecuteSql sSql
   
   sSql = "UPDATE ##Fait SET FA_ITNUMBER='" & sNewRpt & "',FA_ITREVISION='" _
          & sNewRev & "',FA_ITDIMACT='',FA_ITACCEPTED='N'"
   clsADOCon.ExecuteSql sSql
   
   sSql = "INSERT INTO FaitTable SELECT * FROM ##Fait"
   clsADOCon.ExecuteSql sSql
   prg1.Value = 50
   Sleep 200
   
   'FadcTable
   sSql = "SELECT * INTO ##Fadc from FadcTable where FA_DOCNUMBER='" & sOldRpt & "' " _
          & "AND FA_DOCREVISION='" & sOldRev & "'"
   clsADOCon.ExecuteSql sSql
   
   sSql = "UPDATE ##Fadc SET FA_DOCNUMBER='" & sNewRpt & "',FA_DOCREVISION='" _
          & sNewRev & "',FA_DOCCREATED='" & vDate & "'"
   clsADOCon.ExecuteSql sSql
   
   sSql = "INSERT INTO FadcTable SELECT * FROM ##Fadc"
   clsADOCon.ExecuteSql sSql
   prg1.Value = 80
   Sleep 200
   prg1.Value = 100
   MouseCursor 0
   If clsADOCon.ADOErrNum = 0 Then
      clsADOCon.CommitTrans
      MsgBox "First Article Report Successfully Copied.", _
         vbInformation, Caption
      cmbPrt.AddItem txtNew
      txtNew = ""
      txtRev = ""
   Else
      clsADOCon.RollbackTrans
      clsADOCon.ADOErrNum = 0
      MsgBox "Could Not Copy The First Article Report.", _
         vbExclamation, Caption
   End If
   Err.Clear
   sSql = "DROP TABLE ##Fahd"
   clsADOCon.ExecuteSql sSql
   
   sSql = "DROP TABLE ##Fait"
   clsADOCon.ExecuteSql sSql
   
   sSql = "DROP TABLE ##Fadc"
   clsADOCon.ExecuteSql sSql
   
   cmdCpy.Enabled = True
   prg1.Visible = False
   Set RdoCpy = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "copyreport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub txtNew_LostFocus()
   txtNew = CheckLen(txtNew, 30)
   
End Sub


Private Sub txtRev_LostFocus()
   txtRev = Compress(txtRev)
   txtRev = CheckLen(txtRev, 6)
   
End Sub
