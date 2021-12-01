VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form ShopSHf03a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Update MO Routings"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6345
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   6345
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSch 
      Caption         =   "&Schedule"
      Height          =   315
      Left            =   5280
      TabIndex        =   17
      TabStop         =   0   'False
      ToolTipText     =   "Update Entries and Re-Schedule"
      Top             =   2520
      Width           =   875
   End
   Begin VB.CheckBox optHeader 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   252
      Left            =   1440
      TabIndex        =   2
      ToolTipText     =   "Update Only The Rout By And Approved Information"
      Top             =   1800
      Width           =   852
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "ShopSHf03a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdUpd 
      Caption         =   "&Apply"
      Height          =   315
      Left            =   5280
      TabIndex        =   3
      ToolTipText     =   "To Update This MO To The Selected Routing"
      Top             =   600
      Width           =   875
   End
   Begin VB.ComboBox cmbRun 
      Height          =   315
      Left            =   1440
      TabIndex        =   1
      Tag             =   "1"
      ToolTipText     =   "Contains Qualified Runs"
      Top             =   1440
      Width           =   1095
   End
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   1440
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Contains Qualified Part Numbers (Not CO or CL)"
      Top             =   720
      Width           =   3545
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5280
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   90
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   4920
      Top             =   3000
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3210
      FormDesignWidth =   6345
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Header Only"
      ForeColor       =   &H00000000&
      Height          =   252
      Index           =   4
      Left            =   120
      TabIndex        =   16
      Top             =   1800
      Width           =   1212
   End
   Begin VB.Label lblType 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   1440
      TabIndex        =   14
      Top             =   2520
      Width           =   3252
   End
   Begin VB.Label lblRte 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   1440
      TabIndex        =   13
      Top             =   2160
      Width           =   3252
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Type:"
      Height          =   255
      Index           =   8
      Left            =   4920
      TabIndex        =   12
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label lblTyp 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5760
      TabIndex        =   11
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Routing"
      Height          =   252
      Index           =   1
      Left            =   120
      TabIndex        =   10
      Top             =   2160
      Width           =   1332
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      Height          =   255
      Index           =   0
      Left            =   2640
      TabIndex        =   9
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label lblStat 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3360
      TabIndex        =   8
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Run"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1440
      TabIndex        =   6
      Top             =   1080
      Width           =   3135
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   5
      Top             =   765
      Width           =   1335
   End
End
Attribute VB_Name = "ShopSHf03a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'3/21/07 7.3.0 Added Added Routing Columns
Option Explicit
Dim AdoQry As ADODB.Command
Dim AdoParameter As ADODB.Parameter

Dim bCancel As Boolean
Dim bOnLoad As Byte
Dim bGoodMo As Byte
Dim bGoodRun As Byte
Dim sRouting As String

'3/20/07
'Routings
Dim sRtNumber As String
Dim sRtDesc As String
Dim sRtBy As String
Dim sRtAppBy As String
Dim sRtAppDate As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbPrt_Click()
   cmbPrt = GetCurrentPart(cmbPrt, lblDsc)
   lblTyp = GetType()
   GetRouting
   GetRuns
   
End Sub


Private Sub cmbPrt_LostFocus()
   cmbPrt = CheckLen(cmbPrt, 30)
   If bCancel Then Exit Sub
   If Len(Trim(cmbPrt)) > 0 Then
      cmbPrt = GetCurrentPart(cmbPrt, lblDsc)
      lblTyp = GetType()
      GetRouting
      GetRuns
   End If
   
End Sub


Private Sub cmbRun_Click()
   bGoodRun = GetCurrRun()
   
End Sub


Private Sub cmbRun_LostFocus()
   cmbRun = CheckLen(cmbRun, 5)
   If bCancel Then Exit Sub
   cmbRun = Format(Abs(Val(cmbRun)), "####0")
   If Val(cmbRun) > 32767 Then cmbRun = "32767"
   bGoodRun = GetCurrRun()
   
End Sub


Private Sub cmdCan_Click()
   bCancel = True
   Unload Me
   
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   bCancel = True
   
End Sub


Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 4152
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub


Private Sub cmdSch_Click()
   Dim strPassedMo As String
   Dim lRunno As Long
   Dim dtSched As Date
   Dim lQty  As Long
   Dim ret As Integer

   strPassedMo = Compress(Trim(cmbPrt))
   lRunno = cmbRun
   
   ret = GetMoQty(dtSched, lQty)
   If (ret = 0) Then
      MsgBox ("Qty and Schedule not found:" & strPassedMo)
   Else
      Dim mo As New ClassMO
      mo.ScheduleOperations strPassedMo, lRunno, CCur(lQty), dtSched, False
      Set mo = Nothing
      
      MsgBox ("Rescheduled MO: " & strPassedMo & " Run: " & CStr(lRunno))
      
   End If
   
End Sub

   

Private Sub cmdUpd_Click()
   If Trim(sRouting) = "" Then
      MsgBox "Requires A Valid Routing.", _
         vbExclamation, Caption
      Exit Sub
   End If
   
   If bGoodRun = 0 Then
      MsgBox "Requires A Valid Run. See Help.", _
         vbExclamation, Caption
   Else
      bGoodMo = CheckMO()
      If bGoodMo = 0 Then
         If optHeader.Value = vbUnchecked Then
            UpDateRouting
         Else
            UpdateHeader
         End If
      End If
   End If
   
End Sub

Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      FillCombo
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST
   FormatControls
   sSql = "SELECT RUNREF,RUNNO FROM RunsTable WHERE " _
          & "RUNREF= ? AND RUNSTATUS NOT LIKE 'C%' "

   Set AdoQry = New ADODB.Command
   AdoQry.CommandText = sSql
   
   Set AdoParameter = New ADODB.Parameter
   AdoParameter.Type = adChar
   AdoParameter.SIZE = 30
   
   AdoQry.Parameters.Append AdoParameter
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set AdoParameter = Nothing
   Set AdoQry = Nothing
   Set ShopSHf03a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub FillCombo()
   On Error GoTo DiaErr1
   cmbPrt.Clear
   MouseCursor 13
   sSql = "SELECT DISTINCT PARTREF,PARTNUM,RUNREF FROM PartTable,RunsTable " _
          & " WHERE PARTREF=RUNREF AND RUNSTATUS NOT LIKE 'C%' ORDER BY PARTREF"
   LoadComboBox cmbPrt
   If cmbPrt.ListCount > 0 Then
      cmbPrt = cmbPrt.List(0)
      cmbPrt = GetCurrentPart(cmbPrt, lblDsc, True)
      lblTyp = GetType()
      GetRouting
      GetRuns
   End If
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub GetRuns()
   Dim RdoRns As ADODB.Recordset
   cmbRun.Clear
   MouseCursor 13
   AdoQry.Parameters(0).Value = Compress(cmbPrt)
   bSqlRows = clsADOCon.GetQuerySet(RdoRns, AdoQry, ES_FORWARD)
   If bSqlRows Then
      With RdoRns
         Do Until .EOF
            cmbRun.AddItem Format(!Runno, "####0")
            .MoveNext
         Loop
         ClearResultSet RdoRns
      End With
   End If
   MouseCursor 0
   Set RdoRns = Nothing
   If cmbRun.ListCount > 0 Then
      cmbRun = cmbRun.List(0)
      If GetPreferenceValue("AutoSelectLastRun") = "1" Then cmbRun = cmbRun.List(cmbRun.ListCount - 1)
      bGoodRun = GetCurrRun()
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "getruns"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Function GetMoQty(ByRef dtSched As Date, ByRef lQty As Long) As Byte
   
   Dim lRunno As Long
   Dim sPart As String
   
   Dim RdoRun As ADODB.Recordset
   
   lRunno = Val(cmbRun)
   sPart = Compress(cmbPrt)
   'On Error GoTo DiaErr1
   sSql = "SELECT RUNREF,RUNNO,RUNQTY,RUNSCHED FROM RunsTable " _
          & "WHERE RUNREF='" & sPart & "' AND RUNNO=" & lRunno & " "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRun, ES_FORWARD)
   If bSqlRows Then
      With RdoRun
         lQty = Trim(!RUNQTY)
         dtSched = "" & Format(!RUNSCHED, "mm/dd/yy")
         ClearResultSet RdoRun
      End With
      GetMoQty = 1
   Else
      GetMoQty = 0
   End If
   Set RdoRun = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getcurrrun"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function


Private Function GetCurrRun() As Byte
   Dim RdoRun As ADODB.Recordset
   Dim lRunno As Long
   Dim sPart As String
   
   lRunno = Val(cmbRun)
   sPart = Compress(cmbPrt)
   'On Error GoTo DiaErr1
   sSql = "SELECT RUNREF,RUNNO,RUNSTATUS FROM RunsTable " _
          & "WHERE RUNREF='" & sPart & "' AND RUNNO=" & lRunno & " "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRun, ES_FORWARD)
   If bSqlRows Then
      With RdoRun
         lblStat = "" & Trim(!RUNSTATUS)
         ClearResultSet RdoRun
      End With
   Else
      lblStat = "**"
   End If
   If Left(lblStat, 1) <> "C" Then
      GetCurrRun = 1
   Else
      GetCurrRun = 0
      lblStat = "**"
   End If
   Set RdoRun = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getcurrrun"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub lblDsc_Change()
   If Left(lblDsc, 8) = "*** Part" Then
      lblDsc.ForeColor = ES_RED
   Else
      lblDsc.ForeColor = Es_TextForeColor
   End If
   
End Sub

Private Sub lblStat_Change()
   If lblStat = "**" Then
      lblStat.ForeColor = ES_RED
   Else
      lblStat.ForeColor = Es_TextForeColor
   End If
   
End Sub


Private Function CheckMO() As Byte
   Dim RdoOps As ADODB.Recordset
   Dim bByte As Byte
   Dim lRunno As Long
   Dim sPart As String
   Dim sMsg As String
   
   sPart = Compress(cmbPrt)
   lRunno = Val(cmbRun)
   'On Error GoTo DiaErr1
   sSql = "SELECT DISTINCT OPREF,OPRUN,OPCOMPLETE FROM " _
          & "RnopTable WHERE OPCOMPLETE=1 AND (OPREF='" _
          & sPart & "' AND OPRUN=" & lRunno & ") "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoOps, ES_FORWARD)
   If bSqlRows Then bByte = 1 Else bByte = 0
   If bByte = 0 Then
      ' Added PITYPE 16 (cancelled) - so that the MO routing can be routed.
      sSql = "SELECT DISTINCT PIRUNPART,PIRUNNO FROM " _
             & "PoitTable WHERE PIRUNPART='" & sPart & "' " _
             & "AND PIRUNNO=" & lRunno & " AND PITYPE <> 16"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoOps, ES_FORWARD)
      If bSqlRows Then
         sMsg = "This MO Has A Purchase ORDER Allocated." & vbCr _
                & "Remove The Allocation(s) And Retry."
         MsgBox sMsg, vbExclamation, Caption
         bByte = 1
      Else
         bByte = 0
      End If
   Else
      sMsg = "This MO Has Completed Operation(s)." & vbCr _
             & "Uncomplete The Operations And Retry."
      MsgBox sMsg, vbExclamation, Caption
   End If
   
   Set RdoOps = Nothing
   CheckMO = bByte
   Exit Function
   
DiaErr1:
   sProcName = "checkmo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub GetRouting()
   Dim RdoRte As ADODB.Recordset
   Dim sRoutType As String
   ' On Error GoTo DiaErr1
   MouseCursor 13
   sSql = "SELECT PARTREF,PAROUTING,PALEVEL,RTREF,RTNUM FROM PartTable," _
          & "RthdTable WHERE (PARTREF='" & Compress(cmbPrt) & "' " _
          & "AND PAROUTING=RTREF)"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRte, ES_FORWARD)
   If bSqlRows Then
      With RdoRte
         lblRte = "" & Trim(!RTNUM)
         lblType = "Assigned Routing."
         lblTyp = !PALEVEL
         ClearResultSet RdoRte
      End With
   Else
      
      sRoutType = "RTEPART" & Trim(lblTyp)
      sSql = "SELECT " & sRoutType & ",RTREF,RTNUM FROM ComnTable," _
             & "RthdTable WHERE (COREF=1 AND " & sRoutType & "=RTREF)"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoRte, ES_FORWARD)
      If bSqlRows Then
         With RdoRte
            lblRte = "" & Trim(RdoRte.Fields(2))
            lblType = "Default Routing."
            ClearResultSet RdoRte
         End With
      Else
         lblRte = ""
      End If
   End If
   MouseCursor 0
   If Trim(lblRte) = "" Or Left(lblRte, 5) = "No Ro" Then
      lblRte = "No Routing Assignment"
      lblType = "*** Requires A Routing ***"
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "getrouting"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub lblType_Change()
   If Left(lblType, 3) = "***" Then
      lblType.ForeColor = ES_RED
      sRouting = ""
   Else
      lblType.ForeColor = Es_TextForeColor
      sRouting = Compress(lblRte)
   End If
   
End Sub


Private Sub UpDateRouting()
   Dim RdoRte As ADODB.Recordset
   Dim bResponse As Byte
   Dim iCurrOp As Integer
   Dim lRunno As Long
   Dim sMsg As String
   Dim sPart As String
   Dim sRoutCmt As String
   
   sRouting = Compress(lblRte)
   On Error GoTo DiaErr1
   
   sRtNumber = ""
   sRtDesc = ""
   sRtBy = ""
   sRtAppBy = ""
   sRtAppDate = ""
   
   sMsg = "This Replaces The Current Routing (If One)." & vbCr _
          & "Do You Want To Continue With The New One?"
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      sPart = Compress(cmbPrt)
      lRunno = Val(cmbRun)
      'delete existing
      sSql = "DELETE FROM RnopTable WHERE OPREF='" & sPart _
             & "' AND OPRUN=" & lRunno & " "
      clsADOCon.ExecuteSQL sSql
      
      'insert new
      sSql = "SELECT OPREF,OPNO,OPSHOP,OPCENTER,OPSETUP,OPUNIT," _
             & "OPPICKOP,OPSERVPART,OPQHRS,OPMHRS,OPSVCUNIT,OPTOOLLIST,OPCOMT FROM " _
             & "RtopTable WHERE OPREF='" & Compress(sRouting) & "' "
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoRte, ES_KEYSET)
      If bSqlRows Then
         With RdoRte
            Do Until .EOF
               If iCurrOp = 0 Then iCurrOp = !opNo
               On Error Resume Next
               sRoutCmt = "" & Trim(!OPCOMT)
               sRoutCmt = ReplaceString(sRoutCmt)
               sSql = "INSERT INTO RnopTable (OPREF,OPRUN,OPNO,OPSHOP,OPCENTER," _
                      & "OPQHRS,OPMHRS,OPPICKOP,OPSERVPART,OPSUHRS,OPUNITHRS,OPSVCUNIT,OPTOOLLIST,OPCOMT) " _
                      & "VALUES('" & Compress(cmbPrt) & "'," _
                      & Trim(cmbRun) & "," _
                      & !opNo & ",'" _
                      & Trim(!OPSHOP) & "','" _
                      & Trim(!OPCENTER) & "'," _
                      & !OPQHRS & "," _
                      & !OPMHRS & "," _
                      & !OPPICKOP & ",'" _
                      & Trim(!OPSERVPART) & "'," _
                      & !OPSETUP & "," _
                      & !OPUNIT & "," _
                      & !OPSVCUNIT & ",'" _
                      & Trim(!OPTOOLLIST) & "','" _
                      & Trim(sRoutCmt) & "')"
               clsADOCon.ExecuteSQL sSql
               .MoveNext
            Loop
            ClearResultSet RdoRte
         End With
         sSql = "SELECT * FROM RthdTable WHERE RTREF='" & Compress(lblRte) & "' "
         bSqlRows = clsADOCon.GetDataSet(sSql, RdoRte, ES_FORWARD)
         If bSqlRows Then
            With RdoRte
               sRtNumber = "" & Trim(!RTNUM)
               sRtDesc = "" & Trim(!RTDESC)
               sRtBy = "" & Trim(!RTBY)
               sRtAppBy = "" & Trim(!RTAPPBY)
               If Not IsNull(!RTAPPDATE) Then
                  sRtAppDate = Format$(!RTAPPDATE, "mm/dd/yy")
               Else
                  sRtAppDate = ""
               End If
               .Cancel
            End With
         End If
         sSql = "UPDATE RunsTable SET RUNOPCUR=" & iCurrOp & "," _
                & "RUNRTNUM='" & sRtNumber & "'," _
                & "RUNRTDESC='" & sRtDesc & "'," _
                & "RUNRTBY='" & sRtBy & "'," _
                & "RUNRTAPPBY='" & sRtAppBy & "'," _
                & "RUNRTAPPDATE='" & sRtAppDate & "' " _
                & "WHERE RUNREF='" & sPart & "' AND RUNNO=" _
                & lRunno & " "
         clsADOCon.ExecuteSQL sSql
      End If
      MsgBox "Routing Was Successfully Updated.", _
         vbInformation, Caption
   Else
      CancelTrans
   End If
   Exit Sub
   
DiaErr1:
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   
End Sub

Private Function GetType() As Byte
   On Error Resume Next
   Dim RdoTyp As ADODB.Recordset
   sSql = "SELECT PARTREF,PADESC,PALEVEL FROM PartTable Where " _
          & "PARTREF='" & Compress(cmbPrt) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoTyp, ES_FORWARD)
   If bSqlRows Then GetType = RdoTyp!PALEVEL Else GetType = 1
   If bSqlRows Then lblDsc = RdoTyp!PADESC Else lblDsc = ""
   Set RdoTyp = Nothing
   
End Function

'3/21/07

Private Sub UpdateHeader()
   Dim RdoRte As ADODB.Recordset
   Dim bResponse As Byte
   
   bResponse = MsgBox("You Have Checked The Header Only Button. " & vbCr _
               & "Would You Like To Read The Help First?", _
               ES_NOQUESTION, Caption)
   If bResponse = vbYes Then
      CancelTrans
      Exit Sub
   End If
   
   sSql = "SELECT * FROM RthdTable WHERE RTREF='" & Compress(lblRte) & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRte, ES_FORWARD)
   If bSqlRows Then
      With RdoRte
         sRtNumber = "" & Trim(!RTNUM)
         sRtDesc = "" & Trim(!RTDESC)
         sRtBy = "" & Trim(!RTBY)
         sRtAppBy = "" & Trim(!RTAPPBY)
         If Not IsNull(!RTAPPDATE) Then
            sRtAppDate = Format$(!RTAPPDATE, "mm/dd/yy")
         Else
            sRtAppDate = ""
         End If
         .Cancel
      End With
      sSql = "UPDATE RunsTable SET " _
             & "RUNRTNUM='" & sRtNumber & "'," _
             & "RUNRTDESC='" & sRtDesc & "'," _
             & "RUNRTBY='" & sRtBy & "'," _
             & "RUNRTAPPBY='" & sRtAppBy & "'," _
             & "RUNRTAPPDATE='" & sRtAppDate & "' " _
             & "WHERE RUNREF='" & Compress(cmbPrt) & "' AND RUNNO=" _
             & Val(cmbRun) & " "
      clsADOCon.ExecuteSQL sSql
   End If
   MsgBox "Routing Was Successfully Updated.", _
      vbInformation, Caption
   
End Sub
