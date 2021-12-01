VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form FainFAe02a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Revise A First Article Report"
   ClientHeight    =   2355
   ClientLeft      =   2685
   ClientTop       =   1425
   ClientWidth     =   6000
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   Icon            =   "FainFAe02a.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2355
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "FainFAe02a.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox cmbPrt 
      DataSource      =   "rDt1"
      Height          =   315
      Left            =   1560
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Enter a New Part or Select From List (30 chars)"
      Top             =   840
      Width           =   3255
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5760
      Top             =   1920
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2355
      FormDesignWidth =   6000
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "R&evise"
      Height          =   315
      Index           =   1
      Left            =   5040
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Open The Report For Revisions"
      Top             =   600
      Width           =   875
   End
   Begin VB.ComboBox cmbRev 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1560
      Sorted          =   -1  'True
      TabIndex        =   1
      Tag             =   "8"
      Text            =   " "
      ToolTipText     =   "SelectRevision From List"
      Top             =   1560
      Width           =   945
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Index           =   0
      Left            =   5040
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   90
      Width           =   875
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1560
      TabIndex        =   6
      Top             =   1200
      Width           =   3015
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
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Report Revision"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   1560
      Width           =   1545
   End
End
Attribute VB_Name = "FainFAe02a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'New 4/23/03
'12/2/05 Corrected FormUnload
Option Explicit
Dim bGoodReport As Byte
Dim bCancel As Byte
Dim bOnLoad As Byte
Dim bRevise As Byte

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
      '1/25/2009 Should be iList, rather than hardcoded to 1
      'If cmbRev = cmbRev.List(1) Then b = 1
      If cmbRev = cmbRev.List(iList) Then b = 1
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


Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 6202
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub cmdOk_Click(Index As Integer)
   If Len(Trim(cmbPrt)) > 4 Then
      bGoodReport = CheckReport()
   Else
      MsgBox "Your Report Name Must Be 4-30 Characters.", _
         vbInformation, Caption
   End If
   If bGoodReport Then
      bRevise = 1
      FainFAe02b.Show
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
   If bRevise = 0 Then FormUnload
   Set FainFAe02a = Nothing
   
End Sub







Private Function FillCombo() As Byte
   On Error GoTo DiaErr1
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
         GetReport = 1
         ClearResultSet RdoRep
      End With
      FillRevisions
   Else
      GetReport = 0
      lblDsc = "*** Report Wasn't Found ***"
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
   If cmbRev.ListCount > 0 Then cmbRev = cmbRev.List(0)
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
          & "(FA_REF='" & Compress(cmbPrt) & "' AND FA_REVISION='" _
          & Trim(cmbRev) & "')"
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
