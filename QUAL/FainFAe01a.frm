VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form FainFAe01a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New First Article Report"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6375
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "FainFAe01a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CheckBox optVew 
      Caption         =   "View"
      Height          =   255
      Left            =   0
      TabIndex        =   12
      Top             =   360
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdFnd 
      DownPicture     =   "FainFAe01a.frx":07AE
      Height          =   315
      Left            =   4920
      Picture         =   "FainFAe01a.frx":0AF0
      Style           =   1  'Graphical
      TabIndex        =   11
      TabStop         =   0   'False
      ToolTipText     =   "Find A Part Number"
      Top             =   840
      Width           =   350
   End
   Begin VB.TextBox txtRev 
      Height          =   285
      Left            =   1800
      TabIndex        =   2
      Tag             =   "3"
      ToolTipText     =   "Report Revision.  Not Necessarily The Part Revision"
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Create"
      Height          =   315
      Index           =   1
      Left            =   5400
      TabIndex        =   10
      TabStop         =   0   'False
      ToolTipText     =   "Build And Open The Report"
      Top             =   600
      Width           =   875
   End
   Begin VB.TextBox txtLne 
      Height          =   285
      Left            =   4080
      TabIndex        =   3
      Tag             =   "1"
      ToolTipText     =   "Number Of Lines Required (You May Add Or Subtract Later"
      Top             =   1560
      Width           =   615
   End
   Begin VB.TextBox txtDsc 
      Height          =   285
      Left            =   1800
      MaxLength       =   30
      TabIndex        =   1
      Tag             =   "2"
      ToolTipText     =   "Up To 30 Chars"
      Top             =   1200
      Width           =   3015
   End
   Begin VB.TextBox txtPrt 
      Height          =   285
      Left            =   1800
      MaxLength       =   30
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Can, But Needn't Be, A Part Number (30 Char Max)"
      Top             =   840
      Width           =   3015
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5400
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5760
      Top             =   2640
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2550
      FormDesignWidth =   6375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(May Be Altered)"
      Height          =   285
      Index           =   4
      Left            =   4800
      TabIndex        =   9
      Top             =   1560
      Width           =   1635
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Number Of Lines"
      Height          =   285
      Index           =   3
      Left            =   2760
      TabIndex        =   8
      Top             =   1560
      Width           =   1515
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   7
      Top             =   1200
      Width           =   1545
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Report Revision"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   6
      Top             =   1560
      Width           =   1545
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Report Number"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   840
      Width           =   1545
   End
End
Attribute VB_Name = "FainFAe01a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
'12/2/05 Added bShow to bypass FormUnload
Option Explicit
Dim bOnLoad As Byte
Dim bShow As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd


Private Sub cmdCan_Click()
   Unload Me
   
End Sub



Private Sub cmdFnd_Click()
   optVew.Value = vbChecked
   ViewParts.lblControl = "TXTPRT"
   ViewParts.txtPrt = txtPrt
   ViewParts.Show
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 6201
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub cmdOk_Click(Index As Integer)
   Dim b As Byte
   Dim bResponse As Byte
   Dim sMsg As String
   On Error Resume Next
   If Trim(txtPrt) = "" Then
      MsgBox "Requires A Valid Report Number.", _
         vbInformation, Caption
      txtPrt.SetFocus
      Exit Sub
   End If
   If Len(Trim(txtPrt)) < 6 Then
      MsgBox "Requires A Report Number With 6 Or More Chars.", _
         vbInformation, Caption
      txtPrt.SetFocus
      Exit Sub
   End If
   b = IllegalCharacters(txtPrt)
   If b > 0 Then
      MsgBox "The Report Number Contains An Illegal " & Chr$(b) & ".", _
         vbExclamation, Caption
      Exit Sub
   End If
   If Trim(txtDsc) = "" Then
      MsgBox "Requires A Valid Report Description.", _
         vbInformation, Caption
      txtDsc.SetFocus
      Exit Sub
   End If
   If Trim(txtPrt) = "ALL" Then
      MsgBox "ALL Is Not Valid Report Number.", _
         vbInformation, Caption
      txtPrt.SetFocus
      Exit Sub
   End If
   b = CheckReport
   If b = 0 Then
      sMsg = "Do You Want To Create A First Article " & vbCr _
             & "Report For " & txtPrt & "?"
      bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
      If bResponse = vbYes Then
         AddFAInspection
      Else
         CancelTrans
      End If
   Else
      MsgBox "That Report Number And Revision Has Been Recorded.", _
         vbInformation, Caption
   End If
   
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   GetOptions
   FormatControls
   bOnLoad = 1
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveOptions
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   If bShow = 0 Then FormUnload
   Set FainFAe01a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   If Trim(txtLne) = "" Then txtLne = "50"
   
End Sub

Private Sub optVew_Click()
   If optVew.Value = vbUnchecked Then GetDescription
   
End Sub

Private Sub txtDsc_LostFocus()
   txtDsc = CheckLen(txtDsc, 30)
   txtDsc = StrCase(txtDsc)
   
End Sub


Private Sub txtLne_LostFocus()
   txtLne = CheckLen(txtLne, 3)
   txtLne = Format(Abs(Val(txtLne)), "##0")
   If Val(txtLne) = 0 Then txtLne = "50"
   
End Sub


Private Sub txtPrt_Click()
   GetDescription
   
End Sub

Private Sub txtPrt_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF4 Then
      ViewParts.lblControl = "TXTPRT"
      ViewParts.txtPrt = txtPrt
      optVew.Value = vbChecked
      ViewParts.Show
   End If
   
End Sub


Private Sub txtPrt_LostFocus()
   txtPrt = CheckLen(txtPrt, 30)
   GetDescription
   
End Sub



Private Sub GetDescription()
   Dim RdoDsc As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT PARTREF,PADESC FROM PartTable WHERE PARTREF='" _
          & Compress(txtPrt) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoDsc, ES_FORWARD)
   If bSqlRows Then
      With RdoDsc
         txtDsc = "" & Trim(!PADESC)
         ClearResultSet RdoDsc
      End With
   End If
   Set RdoDsc = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getdescr"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub txtRev_LostFocus()
   txtRev = Compress(txtRev)
   txtRev = CheckLen(txtRev, 6)
   
End Sub



Private Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiQual", "Fanew", Trim(sOptions))
   txtLne = sOptions
   
End Sub

Private Sub SaveOptions()
   SaveSetting "Esi2000", "EsiQual", "Fanew", txtLne
   
End Sub


Private Function CheckReport() As Byte
   Dim RdoChk As ADODB.Recordset
   sSql = "SELECT FA_REF,FA_REVISION FROM FahdTable WHERE (" _
          & "FA_NUMBER='" & Compress(txtPrt) & "' AND " _
          & "FA_REVISION='" & Compress(txtRev) & "')"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoChk, ES_FORWARD)
   If bSqlRows Then
      CheckReport = 1
      ClearResultSet RdoChk
   Else
      CheckReport = 0
   End If
   Set RdoChk = Nothing
   
End Function

Private Sub AddFAInspection()
   Dim bByte As Byte
   Dim iRow As Integer
   Dim iLines As Integer
   Dim sReptRef As String
   
   bByte = IllegalCharacters(txtPrt)
   If bByte > 0 Then
      MsgBox "The First Article Number Contains An Illegal " & Chr$(bByte) & ".", _
         vbExclamation, Caption
      Exit Sub
   End If
   sReptRef = Compress(txtPrt)
   iLines = Val(txtLne)
   If iLines = 0 Then iLines = 10
   On Error Resume Next
   'Header
   clsADOCon.ADOErrNum = 0
   clsADOCon.BeginTrans
   sSql = "INSERT INTO FahdTable (FA_REF,FA_NUMBER,FA_REVISION,FA_DESCRIPTION) " _
          & "VALUES ('" & sReptRef & "','" & txtPrt & "','" _
          & txtRev & "','" & txtDsc & "')"
   clsADOCon.ExecuteSQL sSql
   'Detail
   For iRow = 1 To iLines
      sSql = "INSERT INTO FaitTable (FA_ITNUMBER,FA_ITREVISION," _
             & "FA_ITFEATURENUM,FA_ITSEQUENCE) VALUES('" & sReptRef _
             & "','" & txtRev & "'," & iRow & "," & iRow & ")"
      clsADOCon.ExecuteSQL sSql
   Next
   'Documents
   For iRow = 1 To 10
      sSql = "INSERT INTO FadcTable (FA_DOCNUMBER,FA_DOCREVISION," _
             & "FA_DOCITEM) VALUES('" & sReptRef _
             & "','" & txtRev & "'," & iRow & ")"
      clsADOCon.ExecuteSQL sSql
   Next
   If clsADOCon.ADOErrNum = 0 Then
      clsADOCon.CommitTrans
      SysMsg "First Article Report Created", True
      bShow = 1
      FainFAe02b.optFrom.Value = vbChecked
      FainFAe02b.Show
   Else
      clsADOCon.RollbackTrans
      MsgBox "Could Not Create The First Article Report.", _
         vbExclamation, Caption
   End If
   
End Sub
