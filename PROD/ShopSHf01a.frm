VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form ShopSHf01a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cancel A Manufacturing Order"
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6345
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   6345
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "ShopSHf01a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "&Apply"
      Height          =   315
      Left            =   5280
      TabIndex        =   2
      ToolTipText     =   "Cancel This MO"
      Top             =   600
      Width           =   975
   End
   Begin VB.ComboBox cmbRun 
      Height          =   315
      Left            =   1320
      TabIndex        =   1
      Tag             =   "1"
      ToolTipText     =   "Contains Qualified Runs"
      Top             =   1440
      Width           =   1095
   End
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   1320
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Contains Qualified Part Numbers"
      Top             =   720
      Width           =   3545
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5280
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   975
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5640
      Top             =   1560
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2190
      FormDesignWidth =   6345
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      Height          =   255
      Index           =   0
      Left            =   2520
      TabIndex        =   8
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label lblStat 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3240
      TabIndex        =   7
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Run"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   6
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1320
      TabIndex        =   5
      Top             =   1080
      Width           =   3135
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   4
      Top             =   765
      Width           =   1095
   End
End
Attribute VB_Name = "ShopSHf01a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'8/17/05 Clear boxes after cancel
Option Explicit
Dim rdoQry As ADODB.Command
Dim AdoParameter1 As ADODB.Parameter

Dim bOnLoad As Byte
Dim bGoodRun As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbPrt_Click()
   cmbPrt = GetCurrentPart(cmbPrt, lblDsc)
   GetRuns
   
End Sub


Private Sub cmbPrt_LostFocus()
   cmbPrt = CheckLen(cmbPrt, 30)
   If Len(cmbPrt) Then
      cmbPrt = GetCurrentPart(cmbPrt, lblDsc)
      GetRuns
   End If
   
End Sub


Private Sub cmbRun_Click()
   bGoodRun = GetCurrRun()
   
End Sub


Private Sub cmbRun_LostFocus()
   cmbRun = CheckLen(cmbRun, 5)
   cmbRun = Format(Abs(Val(cmbRun)), "####0")
   If Val(cmbRun) > 32767 Then cmbRun = "32767"
   bGoodRun = GetCurrRun()
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   cmbPrt = ""
   
End Sub


Private Sub cmdDel_Click()
   If bGoodRun = 0 Then
      MsgBox "Requires A Valid Run. See Help.", _
         vbExclamation, Caption
   Else
      CancelRun
   End If
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 4150
      cmdHlp = False
      MouseCursor 0
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
          & "RUNREF= ? AND (RUNSTATUS NOT LIKE " _
          & "'C%' AND RUNSTATUS<>'PP' AND RUNSTATUS<>'PC') "
   Set rdoQry = New ADODB.Command
   rdoQry.CommandText = sSql
   
   Set AdoParameter1 = New ADODB.Parameter
   AdoParameter1.Type = adChar
   AdoParameter1.SIZE = 30
   
   rdoQry.Parameters.Append AdoParameter1
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   On Error Resume Next
   Set AdoParameter1 = Nothing
   Set rdoQry = Nothing
   Set ShopSHf01a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub FillCombo()
   On Error GoTo DiaErr1
   cmbPrt.Clear
   sSql = "SELECT DISTINCT PARTREF,PARTNUM,RUNREF FROM " _
          & "RunsTable,PartTable WHERE PARTREF=RUNREF AND " _
          & "(RUNSTATUS NOT LIKE'C%' AND RUNSTATUS<>'PP' AND " _
          & "RUNSTATUS<>'PC') ORDER BY PARTREF"
   LoadComboBox cmbPrt
   If cmbPrt.ListCount > 0 Then
      cmbPrt = cmbPrt.List(0)
      cmbPrt = GetCurrentPart(cmbPrt, lblDsc, True)
   End If
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
   'rdoQry(0) = Compress(cmbPrt)
   rdoQry.Parameters(0).Value = Compress(cmbPrt)
   
   bSqlRows = clsADOCon.GetQuerySet(RdoRns, rdoQry, ES_FORWARD)
   If bSqlRows Then
      With RdoRns
         Do Until .EOF
            AddComboStr cmbRun.hwnd, Format$(!Runno, "####0")
            .MoveNext
         Loop
         ClearResultSet RdoRns
      End With
   End If
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

Private Function GetCurrRun() As Byte
   Dim RdoRun As ADODB.Recordset
   
   Dim lRunno As Long
   Dim sPart As String
   
   lRunno = Val(cmbRun)
   sPart = Compress(cmbPrt)
   On Error GoTo DiaErr1
   sSql = "SELECT RUNREF,RUNNO,RUNSTATUS FROM RunsTable " _
          & "WHERE RUNREF='" & sPart & "' AND RUNNO=" & lRunno & " "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRun, ES_FORWARD)
   If bSqlRows Then
      lblStat = "" & Trim(RdoRun!RUNSTATUS)
   Else
      lblStat = "**"
   End If
   If lblStat = "SC" Or lblStat = "PL" Or lblStat = "RL" Then
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


Private Sub CancelRun()
   Dim RdoPoi As ADODB.Recordset
   Dim bResponse As Byte
   Dim lRunno As Long
   Dim sMsg As String
   Dim sPart As String
   
   sPart = Compress(cmbPrt)
   lRunno = Val(cmbRun)
   On Error GoTo DiaErr1
   
   ' Check for any MO allocation
   Dim mo As New ClassMO
   mo.LoggingEnabled = False
      
   'determine whether picks for unclosed MOs
   If mo.AreTherePicksForUnclosedMos Then
      MouseCursor ccArrow
      MsgBox "There are still Lot tracked parts allocated for this MOs." & vbCr & "Can't cancel MO."
      Exit Sub
   End If
   
   sMsg = "Any Sales Order Allocations Will Be Removed." & vbCr _
          & "Do You Really Want To Cancel This Run?"
   bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
   If bResponse = vbYes Then
      sSql = "SELECT PIRUNPART,PIRUNNO FROM PoitTable " _
             & "WHERE PIRUNPART='" & sPart & "' AND PIRUNNO=" _
             & lRunno & "  AND PITYPE <> 16"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoPoi, ES_FORWARD)
      If bSqlRows Then
         MsgBox "This MO Has Been Allocated To At Least One " & vbCr _
            & "Purchase Order Item. Remove All Allocations.", _
            vbExclamation, Caption
         Set RdoPoi = Nothing
      Else
         clsADOCon.BeginTrans
         clsADOCon.ADOErrNum = 0
         
         On Error Resume Next
         sSql = "DELETE FROM RnalTable WHERE RAREF='" _
                & sPart & "' AND RARUN=" & lRunno & " "
         clsADOCon.ExecuteSQL sSql
         
         sSql = "UPDATE RnopTable SET OPCOMPLETE=1," _
                & "OPCOMPDATE='" & Format(ES_SYSDATE, "mm/dd/yy") & "'," _
                & "OPYIELD=0 WHERE OPREF='" & sPart & "' " _
                & "AND OPRUN=" & lRunno & " "
         clsADOCon.ExecuteSQL sSql
         
         sSql = "DELETE FROM MopkTable WHERE " _
                & "PKMOPART='" & sPart & "' AND " _
                & "PKMORUN=" & lRunno & " "
         clsADOCon.ExecuteSQL sSql
         
         sSql = "UPDATE RunsTable SET RUNSTATUS='CA'," _
                & "RUNPKPURGED=1,RUNCANCELED='" _
                & Format(ES_SYSDATE, "mm/dd/yy hh:mm") & "'," _
                & "RUNCANCELEDBY='" & sInitials & "' " _
                & "WHERE RUNREF='" & sPart & "' AND RUNNO=" & lRunno & " "
         clsADOCon.ExecuteSQL sSql
         If clsADOCon.ADOErrNum = 0 Then
            clsADOCon.CommitTrans
            lblDsc = ""
            lblStat = ""
            cmbRun.Clear
            MsgBox "The Run Was Canceled.", vbInformation, Caption
            FillCombo
         Else
            clsADOCon.RollbackTrans
            MsgBox "Couldn't Cancel The Run.", vbExclamation, Caption
         End If
      End If
   Else
      CancelTrans
   End If
   Set RdoPoi = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "cancelrun"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub
