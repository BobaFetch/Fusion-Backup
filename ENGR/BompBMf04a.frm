VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form BompBMf04a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Release A Parts List To Production"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7110
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   7110
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "BompBMf04a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   16
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdRel 
      Caption         =   "&Apply"
      Height          =   315
      Left            =   5880
      TabIndex        =   14
      ToolTipText     =   "Release/Unrelease Parts List"
      Top             =   600
      Width           =   870
   End
   Begin VB.CommandButton cmdVew 
      DownPicture     =   "BompBMf04a.frx":07AE
      Height          =   320
      Left            =   4800
      Picture         =   "BompBMf04a.frx":0C88
      Style           =   1  'Graphical
      TabIndex        =   13
      TabStop         =   0   'False
      ToolTipText     =   "Parts List for Part and Revision"
      Top             =   1080
      Width           =   350
   End
   Begin VB.CheckBox optRel 
      Enabled         =   0   'False
      Height          =   255
      Left            =   6480
      TabIndex        =   2
      ToolTipText     =   "Status"
      Top             =   1440
      Width           =   252
   End
   Begin VB.ComboBox cmbRev 
      ForeColor       =   &H00800000&
      Height          =   288
      Left            =   5760
      Sorted          =   -1  'True
      TabIndex        =   1
      ToolTipText     =   "Revision-Select From List"
      Top             =   1080
      Width           =   975
   End
   Begin VB.ComboBox cmbPls 
      Height          =   288
      Left            =   1320
      Sorted          =   -1  'True
      TabIndex        =   0
      ToolTipText     =   "Select Part Number"
      Top             =   1080
      Width           =   3405
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5880
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   915
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6120
      Top             =   2400
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3135
      FormDesignWidth =   7110
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Released"
      Height          =   252
      Index           =   2
      Left            =   5280
      TabIndex        =   15
      ToolTipText     =   "Current Status"
      Top             =   1440
      Width           =   1092
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   1320
      TabIndex        =   12
      Top             =   1440
      Width           =   3132
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Effective"
      Height          =   252
      Index           =   3
      Left            =   240
      TabIndex        =   11
      Top             =   2160
      Width           =   1332
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Obsolete"
      Height          =   252
      Index           =   4
      Left            =   240
      TabIndex        =   10
      Top             =   2520
      Width           =   1332
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Reference"
      Height          =   252
      Index           =   5
      Left            =   240
      TabIndex        =   9
      Top             =   1800
      Width           =   1332
   End
   Begin VB.Label txtObs 
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   1320
      TabIndex        =   8
      Top             =   2520
      Width           =   950
   End
   Begin VB.Label txtEff 
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   1320
      TabIndex        =   7
      Top             =   2160
      Width           =   950
   End
   Begin VB.Label txtRef 
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   1320
      TabIndex        =   6
      Top             =   1800
      Width           =   1272
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Rev:"
      Height          =   252
      Index           =   1
      Left            =   5280
      TabIndex        =   5
      Top             =   1080
      Width           =   492
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Parts List"
      Height          =   252
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   1332
   End
End
Attribute VB_Name = "BompBMf04a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'10/24/05 Added cmdRel to clarify Release/Unrelease
Option Explicit
'Dim RdoPrt As rdoQuery
'Dim RdoBmh As ADODB.Recordset
Dim AdoCmdObj As ADODB.Command
Dim RdoBmh As ADODB.Recordset

Dim bGoodPart As Byte
Dim bGoodOList As Byte
Dim bOnLoad As Byte

Dim sPartNumber As String
Dim sPartBomrev As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub cmbPls_Click()
   GetList
   
End Sub


Private Function GetPartsList() As Byte
   cmbRev = Compress(cmbRev)
   sPartNumber = Compress(cmbPls)
   On Error Resume Next
   RdoBmh.Close
   On Error GoTo DiaErr1
   bOnLoad = 1
   sSql = "SELECT BMHREF,BMHREV,BMHREFERENCE,BMHOBSOLETE,BMHREVDATE," _
          & "BMHEFFECTIVE,BMHRELEASED FROM BmhdTable WHERE BMHREF='" & sPartNumber & "' " _
          & "AND BMHREV='" & Trim(cmbRev) & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoBmh)
   If bSqlRows Then
      With RdoBmh
         txtRef = "" & Trim(!BMHREFERENCE)
         txtEff = "" & Format(!BMHEFFECTIVE, "mm/dd/yyyy")
         txtObs = "" & Format(!BMHOBSOLETE, "mm/dd/yyyy")
         optRel.Value = !BMHRELEASED
      End With
      GetPartsList = True
   Else
      txtRef = ""
      txtEff = ""
      txtObs = ""
      optRel.Value = vbUnchecked
      GetPartsList = False
   End If
   If optRel.Value = vbChecked Then cmdRel.Caption = "Un&release" _
                     Else cmdRel.Caption = "&Release "
   bOnLoad = 0
   Exit Function
   
DiaErr1:
   sProcName = "getpartsl"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub cmbPls_LostFocus()
   cmbPls = CheckLen(cmbPls, 30)
   GetList
   
End Sub


Private Sub cmbRev_Click()
   bGoodOList = GetPartsList()
   
End Sub


Private Sub cmbRev_LostFocus()
   cmbRev = CheckLen(cmbRev, 4)
   bGoodOList = GetPartsList()
   
End Sub

Private Sub cmdCan_Click()
   Unload Me
   
End Sub





Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 3253
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub cmdRel_Click()
   Dim bStatus As Byte
   Dim sPartRev As String
   cmbRev = Compress(cmbRev)
   sPartRev = cmbRev
   On Error GoTo DiaErr1
   clsADOCon.ADOErrNum = 0
   If Not bOnLoad Then
      If optRel.Value = vbChecked Then bStatus = 0 _
                        Else bStatus = 1
      sSql = "UPDATE BmhdTable SET BMHRELEASED=" & bStatus & "," & vbCrLf _
         & "BMHRELEASEDATE='" & Format(GetServerDateTime(), "mm/dd/yy") & "'" & vbCrLf _
         & "WHERE BMHREF='" & sPartNumber & "' " _
         & "AND BMHREV='" & sPartRev & "' "
      clsADOCon.ExecuteSQL sSql ' rdExecDirect
      If clsADOCon.ADOErrNum = 0 Then optRel.Value = bStatus
      Sleep 500
      If clsADOCon.RowsAffected Then
         If optRel.Value = vbChecked Then cmdRel.Caption = "Un&release" _
                           Else cmdRel.Caption = "&Release "
         If optRel.Value = vbChecked Then
            SysMsg "Parts Listed Was Released.", True, Me
         Else
            SysMsg "Parts Listed Was Unreleased.", True, Me
         End If
      Else
         MsgBox "Couldn't Find The Parts List Revision Record.", vbExclamation, Caption
      End If
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "cmdRel_click"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub cmdVew_Click()
   If cmdVew Then
      ViewBomTree.Show
      cmdVew = False
   End If
   
End Sub

Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      FillPartsBelow4 cmbPls
      If cmbPls.ListCount > 0 Then cmbPls = cmbPls.List(0)
      'If cUR.CurrentPart <> "" Then cmbPls = cUR.CurrentPart
      GetList
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub GetList()
   Dim RdoRel As ADODB.Recordset
   cmbRev.Clear
   cmbRev = ""
   sPartNumber = Compress(cmbPls)
   On Error GoTo DiaErr1
   AdoCmdObj.Parameters(0) = sPartNumber
   bSqlRows = clsADOCon.GetQuerySet(RdoRel, AdoCmdObj)
   If bSqlRows Then
      With RdoRel
         lblDsc = "" & Trim(!PADESC)
         cmbRev = "" & Trim(!PABOMREV)
         sPartBomrev = "" & Trim(!PABOMREV)
         ClearResultSet RdoRel
      End With
      'cUR.CurrentPart = cmbPls
      bGoodPart = True
   Else
      lblDsc = ""
      txtRef = ""
      txtEff = ""
      txtObs = ""
      sPartBomrev = ""
      MsgBox "Part Wasn't Found or Is The Wrong Type.", vbExclamation, Caption
      bGoodPart = False
   End If
   If bGoodPart Then
      FillBomhRev sPartNumber
      bGoodOList = GetPartsList()
   End If
   Set RdoRel = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getlist"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST
   FormatControls
   sSql = "SELECT PARTREF,PARTNUM,PADESC,PALEVEL,PABOMREV FROM " _
          & "PartTable WHERE PARTREF= ? AND PALEVEL<4"
          
   Set AdoCmdObj = New ADODB.Command
   AdoCmdObj.CommandText = sSql
   
   Dim prmPrtRef As ADODB.Parameter
   Set prmPrtRef = New ADODB.Parameter
   prmPrtRef.Type = adChar
   prmPrtRef.Size = 30
   AdoCmdObj.Parameters.Append prmPrtRef
   'Set RdoPrt = RdoCon.CreatePreparedStatement("", sSql)
   'TODO: RdoPrt.MaxRows = 1
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   SaveCurrentSelections
   FormUnload
   Set AdoCmdObj = Nothing
   Set RdoBmh = Nothing
   Set BompBMf04a = Nothing
   
End Sub

