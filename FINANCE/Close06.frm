VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Close06 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Recost Completed & Closed Manufacturing Orders"
   ClientHeight    =   8925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13605
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8925
   ScaleWidth      =   13605
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox ChkIgGLPost 
      Height          =   255
      Left            =   2160
      TabIndex        =   31
      ToolTipText     =   "Workstation Setting - Allow To Close Without Testing Expendables (Type 5) "
      Top             =   3960
      Width           =   495
   End
   Begin VB.CommandButton cmdViewMoCost 
      Caption         =   "View MO detail Cost"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   11400
      TabIndex        =   30
      ToolTipText     =   " View Manufacturing Order Cost"
      Top             =   7800
      Width           =   1920
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear Selection"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   11280
      TabIndex        =   29
      ToolTipText     =   " Close this Manufacturing Order"
      Top             =   4320
      Width           =   1920
   End
   Begin VB.CommandButton cmdSelMos 
      Caption         =   "Select MO's"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   5880
      TabIndex        =   28
      ToolTipText     =   " Close this Manufacturing Order"
      Top             =   1560
      Width           =   1800
   End
   Begin VB.Frame fraMO 
      Caption         =   "Or recost an individual MO"
      Height          =   1395
      Left            =   600
      TabIndex        =   22
      Top             =   2460
      Width           =   5055
      Begin VB.ComboBox cmbPrt 
         Height          =   315
         Left            =   1320
         TabIndex        =   6
         Tag             =   "3"
         ToolTipText     =   "Contains Qualified Part Numbers (CO)"
         Top             =   420
         Width           =   3545
      End
      Begin VB.ComboBox cmbRun 
         Height          =   315
         Left            =   1320
         TabIndex        =   7
         Tag             =   "1"
         ToolTipText     =   "Contains Qualified Runs"
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Part Number"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   420
         Width           =   1095
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Run Number"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   23
         Top             =   840
         Width           =   1095
      End
   End
   Begin VB.OptionButton optMO 
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   2820
      Width           =   375
   End
   Begin VB.OptionButton optDateRange 
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   315
   End
   Begin VB.CheckBox chkDiagnose 
      Height          =   255
      Left            =   9720
      TabIndex        =   8
      ToolTipText     =   "Workstation Setting - Allow To Close Without Testing Expendables (Type 5) "
      Top             =   3960
      Width           =   495
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   15
      Top             =   8550
      Width           =   13605
      _ExtentX        =   23998
      _ExtentY        =   661
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "Close06.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton optPrn 
      Height          =   360
      Left            =   5160
      Picture         =   "Close06.frx":07AE
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Print The Report"
      Top             =   120
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdVew 
      DownPicture     =   "Close06.frx":0938
      Height          =   350
      Left            =   4680
      Picture         =   "Close06.frx":0E12
      Style           =   1  'Graphical
      TabIndex        =   11
      TabStop         =   0   'False
      ToolTipText     =   "View Last Closed Run Log (Requires A Text Viewer) "
      Top             =   120
      Width           =   360
   End
   Begin VB.CommandButton cmdCloseMOs 
      Caption         =   "Recost"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   11280
      TabIndex        =   9
      ToolTipText     =   " Close this Manufacturing Order"
      Top             =   5040
      Width           =   1920
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   5880
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   120
      Width           =   1800
   End
   Begin VB.Frame fraDateRange 
      Height          =   1515
      Left            =   600
      TabIndex        =   17
      Top             =   720
      Width           =   5055
      Begin VB.TextBox txtPasses 
         Height          =   285
         Left            =   3000
         MaxLength       =   1
         TabIndex        =   4
         Text            =   "9"
         Top             =   1080
         Width           =   195
      End
      Begin VB.TextBox txtMax 
         Height          =   285
         Left            =   5760
         MaxLength       =   5
         TabIndex        =   3
         Text            =   "99999"
         Top             =   1800
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.ComboBox cmbCompletedThru 
         Height          =   315
         Left            =   3000
         TabIndex        =   2
         Tag             =   "4"
         Top             =   600
         Width           =   1095
      End
      Begin VB.ComboBox cmbCompletedFrom 
         Height          =   315
         Left            =   3000
         TabIndex        =   1
         Tag             =   "4"
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Passes"
         Height          =   255
         Left            =   3360
         TabIndex        =   26
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Stop after"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   1110
         Width           =   2355
      End
      Begin VB.Label Label2 
         Caption         =   "Recost a maximum of"
         Height          =   255
         Left            =   3000
         TabIndex        =   21
         Top             =   1830
         Visible         =   0   'False
         Width           =   2355
      End
      Begin VB.Label Label3 
         Caption         =   "MO's"
         Height          =   255
         Left            =   6420
         TabIndex        =   20
         Top             =   1860
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Through"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   19
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Recost MO's completed from"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   18
         Top             =   300
         Width           =   2175
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Grd 
      Height          =   4215
      Left            =   360
      TabIndex        =   27
      ToolTipText     =   "Click To Select Or Scroll And Press Enter (Also Page Up And Page Down)"
      Top             =   4320
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   7435
      _Version        =   393216
      Rows            =   3
      Cols            =   9
      FixedRows       =   2
      FixedCols       =   0
      FocusRect       =   2
      ScrollBars      =   2
      MergeCells      =   1
   End
   Begin VB.Label z1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Ignore GL posted flag "
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   3
      Left            =   480
      TabIndex        =   32
      ToolTipText     =   "Workstation Setting - Allow To Close Without Testing Expendables (Type 5) "
      Top             =   3960
      Width           =   1635
   End
   Begin VB.Image Chkno 
      Height          =   210
      Left            =   6240
      Picture         =   "Close06.frx":12EC
      Stretch         =   -1  'True
      Top             =   1080
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Chkyes 
      Height          =   210
      Left            =   6240
      Picture         =   "Close06.frx":1676
      Stretch         =   -1  'True
      Top             =   840
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Label z1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Diagnose only (do not update costs)"
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   1
      Left            =   7020
      TabIndex        =   16
      ToolTipText     =   "Workstation Setting - Allow To Close Without Testing Expendables (Type 5) "
      Top             =   3960
      Width           =   2715
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Recost Manufacturing Orders"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   240
      TabIndex        =   14
      Top             =   360
      Width           =   5775
   End
End
Attribute VB_Name = "Close06"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
Option Explicit
'Dim RdoQry As rdoQuery
'Dim cmdObj As ADODB.Command

Dim bCantClose As Byte
Dim bOnLoad As Byte
Dim bGoodPrt As Byte
Dim bGoodRun As Byte
Dim bLotsOn As Byte

Private Const BASE_WHERE_CLAUSE = _
   "where RUNSTATUS <> 'CA'" & vbCrLf _
   & "and exists (select INMOPART from InvaTable where INMOPART = RUNREF and INMORUN = RUNNO)"

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbCompletedFrom_DropDown()
   ShowCalendar Me
End Sub

Private Sub cmbCompletedThru_DropDown()
   ShowCalendar Me
End Sub

Private Sub cmbPrt_Change()
   Dim mo As New ClassMO
'   mo.FillComboBoxWithMoRuns cmbRun, "WHERE RUNSTATUS IN ( 'CO', 'CL' ) AND RUNREF = '" & cmbPrt.Text & "'"
   mo.FillComboBoxWithMoRuns cmbRun, _
      BASE_WHERE_CLAUSE & vbCrLf _
      & "AND RUNREF = '" & Compress(cmbPrt.Text) & "'"
End Sub

Private Sub cmbPrt_Click()
   Dim mo As New ClassMO
'   mo.FillComboBoxWithMoRuns cmbRun, _
'      "WHERE RUNSTATUS IN ( 'CO', 'CL' ) AND RUNREF = '" & Compress(cmbPrt.Text) & "'"
   mo.FillComboBoxWithMoRuns cmbRun, _
      BASE_WHERE_CLAUSE & vbCrLf _
      & "AND RUNREF = '" & Compress(cmbPrt.Text) & "'"
End Sub

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdClear_Click()
    Dim iList As Integer
    For iList = 1 To Grd.Rows - 1
        Grd.Col = 7
        Grd.row = iList
        ' Only if the part is checked
        If Grd.CellPicture = Chkyes.Picture Then
            Set Grd.CellPicture = Chkno.Picture
        End If
    Next
End Sub

Private Sub cmdCloseMOs_Click()
   'Dim RdoQty As ADODB.Recordset
   Dim bByte As Byte
   Dim lClose As Long
   Dim lComplete As Long
   
   bCantClose = 0
   
   cmdCloseMOs.Enabled = False
   cmdCan.Enabled = False
   MouseCursor ccHourglass
   
   
   'get the list of all completed or closed MO's in the date range
   Dim mo As New ClassMO
   mo.LoggingEnabled = True
   Dim rdo As ADODB.Recordset
   Dim success As Boolean

   mo.DiagnoseOnly = CBool(chkDiagnose.Value)
   mo.Log "Recosting manufacturing orders completed between " & cmbCompletedFrom & " and " & cmbCompletedThru
   If mo.DiagnoseOnly Then
      mo.Log "Diagnosing Only.  MOs will not be updated"
   End If
   
   'loop through MOs costing only those that have all contributing items costed
   Dim mosToCostThisPass As Long
   Dim firstpass As Boolean
   Dim done As Boolean
   Dim doFinalPass As Boolean
   Dim passNumber As Integer, maxPasses As Integer
   Dim iList As Integer
   Dim strMoPartRef, strMoRun As String
   
   passNumber = 1
   
   mo.Log ""
   mo.Log "** Pass " & passNumber
            
   Dim recostedMos As Integer
   If optDateRange.Value Then
      For iList = 1 To Grd.Rows - 1
          Grd.Col = 8
          Grd.row = iList
          ' Only if the part is checked
          If Grd.CellPicture = Chkyes.Picture Then
            
            Grd.Col = 0
            strMoPartRef = Trim(Grd.Text)
            Grd.Col = 1
            strMoRun = Trim(Grd.Text)
         
            ' Recost MO
            recostedMos = RecostMO(mo, strMoPartRef, strMoRun)
            recostedMos = recostedMos + 1
            
            StatusBar1.SimpleText = " MO : " & mo.PartNumber & " run " & mo.RunNumber _
               & " completed. "
         End If
      Next
   Else
      strMoPartRef = Trim(Compress(cmbPrt.Text))
      strMoRun = Trim(cmbRun.Text)
      ' Recost MO
      RecostMO mo, strMoPartRef, strMoRun
   End If
   
   Dim sMsg As String
   sMsg = "Recosted " & recostedMos & " MOs."
   mo.Log sMsg
   StatusBar1.SimpleText = sMsg & "  See log"
   cmdCloseMOs.Enabled = True
   cmdCan.Enabled = True
   MouseCursor ccArrow
   Beep

End Sub


Private Function RecostMO(ByRef clsMO As ClassMO, ByVal strMoPartRef As String, ByVal strMoRun As String)
   
   ' MO
   clsMO.PartNumber = strMoPartRef
   clsMO.RunNumber = strMoRun
   
   'clear costed flags for records in selected date range.
   'they will be set as we iterate through the layers of MO's feeding other MO's
   sSql = "update RunsTable set RUNMAINTCOSTED = 0" & vbCrLf _
      & "where RUNREF = '" & strMoPartRef & "'" & vbCrLf _
      & "and RUNNO = '" & strMoRun & "'" & vbCrLf
   Debug.Print sSql
   clsADOCon.ExecuteSQL sSql ' rdExecDirect
   
   sSql = "update InvaTable set INMAINTCOSTED = 0 where INTYPE = 10" & vbCrLf _
      & "and INMOPART = '" & strMoPartRef & "'" & vbCrLf _
      & "and INMORUN = '" & strMoRun & "'" & vbCrLf
   Debug.Print sSql
   clsADOCon.ExecuteSQL sSql ' rdExecDirect
   
   sSql = "update LohdTable set LOTMAINTCOSTED = 0 where LOTMOPARTREF <> ''" & vbCrLf _
      & "and LOTMOPARTREF = '" & strMoPartRef & "'" & vbCrLf _
      & "and LOTMORUNNO = '" & strMoRun & "' " & vbCrLf
   Debug.Print sSql
   clsADOCon.ExecuteSQL sSql ' rdExecDirect
   
'   StatusBar1.SimpleText = "Pass " & passNumber & " MO " & totalMos & " of " _
'      & totalRowsThisPass & ": " & clsMO.PartNumber & " run " & clsMO.RunNumber _
'      & " completed " & !RUNCOMPLETE
   clsMO.Log ""
'   clsMO.Log "Pass " & passNumber & " MO " & totalMos & " of " & totalRowsThisPass _
'      & ": " & clsMO.PartNumber & " run " & mo.RunNumber & " completed " & !RUNCOMPLETE
   
   ' Recost the MO
   If (ChkIgGLPost.Value = vbChecked) Then
      clsMO.RecalCostsIgnoreGLPost
   Else
      clsMO.RecalculateCosts
   End If

End Function

Private Sub old_cmdCloseMOs_Click()
  ' Dim RdoQty As ADODB.Recordset
   Dim bByte As Byte
   Dim lClose As Long
   Dim lComplete As Long
   
   bCantClose = 0
   
   cmdCloseMOs.Enabled = False
   cmdCan.Enabled = False
   MouseCursor ccHourglass
   
   
   'get the list of all completed or closed MO's in the date range
   Dim mo As New ClassMO
   mo.LoggingEnabled = True
   Dim rdo As ADODB.Recordset
   Dim success As Boolean

   mo.DiagnoseOnly = CBool(chkDiagnose.Value)
   mo.Log "Recosting manufacturing orders completed between " & cmbCompletedFrom & " and " & cmbCompletedThru
   If mo.DiagnoseOnly Then
      mo.Log "Diagnosing Only.  MOs will not be updated"
   End If
   
   'loop through MOs costing only those that have all contributing items costed
   Dim mosToCostThisPass As Long
   Dim firstpass As Boolean
   Dim done As Boolean
   Dim doFinalPass As Boolean
   Dim passNumber As Integer, maxPasses As Integer
   passNumber = 0
   maxPasses = CInt(txtPasses)
   
   Do
   
      passNumber = passNumber + 1
      If passNumber > maxPasses Then Exit Do
      
      If optDateRange.Value Then
         mo.Log ""
         mo.Log "** Pass " & passNumber
         
         'clear costed flags for records in selected date range.
         'they will be set as we iterate through the layers of MO's feeding other MO's
         If passNumber = 1 Then
            sSql = "update RunsTable set RUNMAINTCOSTED = 0" & vbCrLf _
               & "where RUNCOMPLETE >= '" & Format(cmbCompletedFrom, "mm/dd/yyyy") & "'" & vbCrLf _
               & "and RUNCOMPLETE <= '" & Format(cmbCompletedThru, "mm/dd/yyyy") & "'" & vbCrLf
            Debug.Print sSql
            clsADOCon.ExecuteSQL sSql ' rdExecDirect
            
            sSql = "update InvaTable set INMAINTCOSTED = 0 where INTYPE = 10" & vbCrLf _
               & "and INADATE >= '" & Format(cmbCompletedFrom, "mm/dd/yyyy") & "'" & vbCrLf _
               & "and INADATE <= '" & Format(cmbCompletedThru, "mm/dd/yyyy") & "'" & vbCrLf
            Debug.Print sSql
            clsADOCon.ExecuteSQL sSql ' rdExecDirect
            
            sSql = "update LohdTable set LOTMAINTCOSTED = 0 where LOTMOPARTREF <> ''" & vbCrLf _
               & "and LOTADATE >= '" & Format(cmbCompletedFrom, "mm/dd/yyyy") & "'" & vbCrLf _
               & "and LOTADATE <= '" & Format(cmbCompletedThru, "mm/dd/yyyy") & "'" & vbCrLf
            Debug.Print sSql
            clsADOCon.ExecuteSQL sSql ' rdExecDirect
            
         End If
         
         
         Dim fromClause As String, totalRowsThisPass As Long
'         fromClause = "from RunsTable" & vbCrLf _
'            & "left join viewLotCostsByMoSummary vl on vl.MoPart = RUNREF and vl.MoRun = RUNNO" & vbCrLf _
'            & "left join viewNonLotCostsByMoSummary vn on vn.MoPart = RUNREF and vn.MoRun = RUNNO" & vbCrLf _
'            & "where RUNSTATUS in ( 'CO', 'CL' )" & vbCrLf _
'            & "and RUNCOMPLETE >= '" & Format(cmbCompletedFrom, "mm/dd/yyyy") & "'" & vbCrLf _
'            & "and RUNCOMPLETE <= '" & Format(cmbCompletedThru, "mm/dd/yyyy") & "'" & vbCrLf _
'            & "and RUNMAINTCOSTED = 0" & vbCrLf
         
'         fromClause = "from RunsTable" & vbCrLf _
'            & "left join viewLotCostsByMoSummary vl on vl.MoPart = RUNREF and vl.MoRun = RUNNO" & vbCrLf _
'            & "left join viewNonLotCostsByMoSummary vn on vn.MoPart = RUNREF and vn.MoRun = RUNNO" & vbCrLf _
'            & "where RUNSTATUS <> 'CA'" & vbCrLf _
'            & "and exists (select INMOPART from InvaTable where INMOPART = RUNREF and INMORUN = RUNNO)" & vbCrLf _
'            & "and RUNCOMPLETE >= '" & Format(cmbCompletedFrom, "mm/dd/yyyy") & "'" & vbCrLf _
'            & "and RUNCOMPLETE <= '" & Format(cmbCompletedThru, "mm/dd/yyyy") & "'" & vbCrLf _
'            & "and RUNMAINTCOSTED = 0" & vbCrLf

         fromClause = "from RunsTable" & vbCrLf _
            & "left join viewLotCostsByMoSummary vl on vl.MoPart = RUNREF and vl.MoRun = RUNNO" & vbCrLf _
            & "left join viewNonLotCostsByMoSummary vn on vn.MoPart = RUNREF and vn.MoRun = RUNNO" & vbCrLf _
            & BASE_WHERE_CLAUSE & vbCrLf _
            & "and RUNCOMPLETE >= '" & Format(cmbCompletedFrom, "mm/dd/yyyy") & "'" & vbCrLf _
            & "and RUNCOMPLETE <= '" & Format(cmbCompletedThru, "mm/dd/yyyy") & "'" & vbCrLf _
            & "and RUNMAINTCOSTED = 0" & vbCrLf

         If Not doFinalPass Then
            fromClause = fromClause & "and isnull( vl.FullyCosted, 1 ) = 1" & vbCrLf _
            & "and isnull( vn.FullyCosted, 1 ) = 1" & vbCrLf
         End If
      Else
'         fromClause = "from RunsTable" & vbCrLf _
'            & "where RUNSTATUS in ( 'CO', 'CL' )" & vbCrLf _
'            & "and RUNREF = '" & Compress(cmbPrt.Text) & "' and RUNNO = " & cmbRun.Text & vbCrLf

'         fromClause = "from RunsTable" & vbCrLf _
'            & "where RUNSTATUS <> 'CA'" & vbCrLf _
'            & "and exists (select INMOPART from InvaTable where INMOPART = RUNREF and INMORUN = RUNNO)" & vbCrLf _
'            & "and RUNREF = '" & Compress(cmbPrt.Text) & "' and RUNNO = " & cmbRun.Text & vbCrLf

         fromClause = "from RunsTable" & vbCrLf _
            & BASE_WHERE_CLAUSE & vbCrLf _
            & "and RUNREF = '" & Compress(cmbPrt.Text) & "' and RUNNO = " & cmbRun.Text & vbCrLf
         done = True
      End If
      
      'first get count of rows to cost this pass
      sSql = "select count(*) " & vbCrLf & fromClause
      If clsADOCon.GetDataSet(sSql, rdo, ES_FORWARD) Then
         totalRowsThisPass = rdo.Fields(0)
      Else
         totalRowsThisPass = 0
      End If
      Set rdo = Nothing
      
      
      sSql = "select rtrim(RUNREF) as RUNREF, RUNNO, RUNCOMPLETE" & vbCrLf _
         & fromClause & "order by RUNREF, RUNNO" & vbCrLf
      
      Dim totalMos As Integer
      Dim recostedMos As Integer
      
      totalMos = 0
      recostedMos = 0
      
      Debug.Print
      Debug.Print "Pass " & passNumber
      Debug.Print sSql
      If clsADOCon.GetDataSet(sSql, rdo, ES_FORWARD) Then
'         cmdCloseMOs.enabled = False
'         cmdCan.enabled = False
'         MouseCursor ccHourglass
         Dim max As Long
         If chkDiagnose.Value = vbChecked Then
            max = 99999
         ElseIf IsNumeric(txtMax.Text) Then
            max = CLng(txtMax.Text)
         Else
            max = 99999
         End If
            
         With rdo
            Do While Not .EOF
               totalMos = totalMos + 1
               mo.PartNumber = !RUNREF
               mo.RunNumber = !Runno
               If totalMos <= max Then
                  StatusBar1.SimpleText = "Pass " & passNumber & " MO " & totalMos & " of " _
                     & totalRowsThisPass & ": " & mo.PartNumber & " run " & mo.RunNumber _
                     & " completed " & !RUNCOMPLETE
                  mo.Log ""
                  mo.Log "Pass " & passNumber & " MO " & totalMos & " of " & totalRowsThisPass _
                     & ": " & mo.PartNumber & " run " & mo.RunNumber & " completed " & !RUNCOMPLETE
                  If mo.RecalculateCosts() Then
                     recostedMos = recostedMos + 1
                  End If
               Else
                  mo.Log "Pass " & passNumber & " MO " & totalMos & " of " & totalRowsThisPass _
                     & ": " & "Did not recost MO # " & totalMos & ": " & mo.PartNumber & " run " & mo.RunNumber
               End If
               .MoveNext
            Loop
         End With
         Set rdo = Nothing
         
         mo.Log ""
         Dim sMsg As String
         If doFinalPass Then
            sMsg = "Final pass (" & passNumber & "): recosted " & recostedMos & " of " & totalMos & " MOs with best available data."
            done = True
         Else
            sMsg = "Pass " & passNumber & ": recosted " & recostedMos & " of " & totalMos & " MOs."
         End If
         mo.Log sMsg
      
      ElseIf optDateRange.Value Then
         If doFinalPass Then
            done = True
         Else
            doFinalPass = True
         End If
      End If
      
   Loop While Not done
   
   StatusBar1.SimpleText = sMsg & "  See log"
   cmdCloseMOs.Enabled = True
   cmdCan.Enabled = True
   MouseCursor ccArrow
   Beep

End Sub


Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 4153
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub


Private Sub cmdSelMos_Click()
   FillGrid
End Sub

Private Sub cmdVew_Click()
   MouseCursor 13
   On Error GoTo DiaErr1
   'SetMdiReportsize MDISect
   
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   sCustomReport = GetCustomReport("closedruns")
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "RequestBy"
   aFormulaValue.Add CStr("'" & sFacility & "'")
   aFormulaValue.Add CStr("'Requested By: " & sInitials & "'")
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
    
   cCRViewer.CRViewerSize Me
   cCRViewer.SetDbTableConnection
   cCRViewer.OpenCrystalReportObject Me, aFormulaName
   cCRViewer.ShowGroupTree False
   
   cCRViewer.ClearFieldCollection aRptPara
   cCRViewer.ClearFieldCollection aFormulaName
   cCRViewer.ClearFieldCollection aFormulaValue
   
'   MDISect.Crw.Formulas(0) = "CompanyName='" & sFacility & "'"
'   MDISect.Crw.Formulas(1) = "RequestBy = 'Requested By: " & sInitials & "'"
'   sCustomReport = GetCustomReport("closedruns")
'   MDISect.Crw.ReportFileName = sReportPath & sCustomReport
'   SetCrystalAction Me
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "printreport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub cmdViewMoCost_Click()
   
   ViewMOCost.txtCompletedFrom = Format(cmbCompletedFrom, "mm/dd/yyyy")
   ViewMOCost.txtCompletedThru = Format(cmbCompletedThru, "mm/dd/yyyy")
   
   ViewMOCost.Show
End Sub

Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      GetSettings
      bLotsOn = CheckLotStatus
      Dim mo As New ClassMO
      mo.FillComboBoxWithMoParts cmbPrt, BASE_WHERE_CLAUSE
      optDateRange_Click
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST
   FormatControls
   
   With Grd
      .ColAlignment(0) = 0
      .ColAlignment(1) = 1
      .ColAlignment(2) = 1
      .ColAlignment(3) = 1
      .ColAlignment(4) = 1
      .ColAlignment(5) = 1
      .ColAlignment(6) = 1
      .ColAlignment(7) = 1
      .ColAlignment(8) = 1
'      .ColAlignment(9) = 1
'      .ColAlignment(10) = 1
'      .ColAlignment(11) = 2
      .Rows = 1
      .row = 0
      .Col = 0
      .Text = "MO PartNumber"
      .Col = 1
      .Text = "MO Run"
      .Col = 2
      .Text = "Lot Number"
      .Col = 3
      .Text = "GL Posted"
      .Col = 4
      .Text = "Qty"
      .Col = 5
      .Text = "Inva UnitCost"
      .Col = 6
      .Text = "Lot UnitCost"
      .Col = 7
      .Text = "ReCost UnitCost"
      .Col = 8
      .Text = "Apply"
      
      .ColWidth(0) = 2300
      .ColWidth(1) = 700
      .ColWidth(2) = 1500
      .ColWidth(3) = 700
      .ColWidth(4) = 1000
      .ColWidth(5) = 1200
      .ColWidth(6) = 1200
      .ColWidth(7) = 1200
      .ColWidth(8) = 700
'      .ColWidth(8) = 1000
'      .ColWidth(9) = 1000
'      .ColWidth(10) = 1500
      .ScrollBars = flexScrollBarBoth
      .AllowUserResizing = flexResizeColumns
      
   End With
   
   bOnLoad = 1
End Sub

Function FillGrid() As Integer
   Dim RdoGrd As ADODB.Recordset
   Dim strEmp As String
   Dim strPartRef As String
   Dim strRunNo As String
   Dim strLotNum As String
   Dim cQty As Currency
   Dim cInAmt As Currency
   Dim cLotUnitCost As Currency
   Dim cReUnitCost  As Currency
   
   MouseCursor ccHourglass
   On Error Resume Next
   Grd.Rows = 1
   On Error GoTo DiaErr1
       
       
   If optDateRange.Value Then
      Dim strComFrom, strComThru As String
      strComFrom = Format(cmbCompletedFrom, "mm/dd/yyyy")
      strComThru = Format(cmbCompletedThru, "mm/dd/yyyy")
      
      
      sSql = "SELECT RUNREF, Runno, LotNumber, INGLPOSTED," & vbCrLf _
               & " INAQTY,  INAMT, LotUnitCost " & vbCrLf _
            & " From RunsTable, INVATABLE, LohdTable, PartTable" & vbCrLf _
            & " Where PartRef = RUNREF And LOTMOPARTREF = RUNREF And LOTMORUNNO = Runno " & vbCrLf _
               & " AND INLOTNUMBER = LotNumber " & vbCrLf _
               & " AND  RUNREF = INMOPART AND RUNNO  = INMORUN" & vbCrLf _
               & " AND RUNCOMPLETE BETWEEN '" & strComFrom & "' AND '" & strComThru & "'" & vbCrLf _
               & " AND RUNSTATUS IN('CL', 'CO')" & vbCrLf _
               & " AND INTYPE = 6" & vbCrLf _
               & " ORDER BY 1"
               '" AND INAMT <> LotUnitCost "

Debug.Print sSql

      bSqlRows = clsADOCon.GetDataSet(sSql, RdoGrd, ES_FORWARD)
      If bSqlRows Then
          With RdoGrd
              Do Until .EOF
              
               strPartRef = Trim(!RUNREF)
               strRunNo = Trim(!Runno)
               strLotNum = Trim(!lotNumber)
               cQty = CDbl(Trim(!INAQTY))
               cInAmt = CDbl(Trim(!INAMT))
               cLotUnitCost = CDbl(Trim(!LotUnitCost))
               cReUnitCost = 0
               FindNewReCostValue strPartRef, strRunNo, cReUnitCost
                              
               If ((Round(cInAmt, 2) <> Round(cReUnitCost, 2)) Or (Round(cReUnitCost, 2) <> Round(cLotUnitCost, 2))) Then
                  Grd.Rows = Grd.Rows + 1
                  Grd.row = Grd.Rows - 1
                  Grd.Col = 0
                  Grd.Text = "" & Trim(!RUNREF)
                  Grd.Col = 1
                  Grd.Text = "" & Trim(!Runno)
                  Grd.Col = 2
                  Grd.Text = "" & Trim(!lotNumber)
                  Grd.Col = 3
                  Grd.Text = "" & Trim(!INGLPOSTED)
                  Grd.Col = 4
                  Grd.Text = "" & Trim(!INAQTY)
                  Grd.Col = 5
                  Grd.Text = "" & Trim(!INAMT)
                  Grd.Col = 6
                  Grd.Text = "" & Trim(!LotUnitCost)
                  Grd.Col = 7
                  Grd.Text = "" & cReUnitCost
                  Grd.Col = 8
                  Set Grd.CellPicture = Chkno.Picture
               End If
               .MoveNext
            Loop
         ClearResultSet RdoGrd
         End With
      End If
      Set RdoGrd = Nothing
   End If
   
   MouseCursor ccArrow
   
   Exit Function
   
DiaErr1:
   sProcName = "fillgrid"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub grd_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeySpace Then
      Grd.Col = 8
      If Grd.row >= 1 Then
         If Grd.row = 0 Then Grd.row = 1
         If Grd.CellPicture = Chkyes.Picture Then
            Set Grd.CellPicture = Chkno.Picture
         Else
            Set Grd.CellPicture = Chkyes.Picture
         End If
      End If
    End If
   

End Sub


Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   SaveSettings
   On Error Resume Next
   FormUnload
   Set Close06 = Nothing
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Private Sub cmbCloseDate_DropDown()
   ShowCalendar Me
End Sub

Private Sub SaveSettings()
   SaveSetting "Esi2000", "EsiFina", "Close06", Trim(str(chkDiagnose)) & "00000"
   SaveSetting "Esi2000", "EsiFina", "Close06.MaxMos", txtMax.Text
   SaveSetting "Esi2000", "EsiFina", "Close06.MaxPasses", txtPasses.Text
   SaveSetting "Esi2000", "EsiFina", "Close06.From", cmbCompletedFrom.Text
   SaveSetting "Esi2000", "EsiFina", "Close06.Thru", cmbCompletedThru.Text
End Sub

Private Sub GetSettings()
   Dim bits As String
   bits = GetSetting("Esi2000", "EsiFina", "Close06", "00000000")
   If Len(bits) < 6 Then bits = "000000"
   chkDiagnose.Value = CInt(Mid(bits, 1, 1))
   txtMax.Text = GetSetting("Esi2000", "EsiFina", "Close06.MaxMos", "99999")
   txtPasses.Text = GetSetting("Esi2000", "EsiFina", "Close06.MaxPasses", "9")
   cmbCompletedFrom.Text = GetSetting("Esi2000", "EsiFina", "Close06.From", Format(ES_SYSDATE, "mm/dd/yy"))
   cmbCompletedThru.Text = GetSetting("Esi2000", "EsiFina", "Close06.Thru", Format(ES_SYSDATE, "mm/dd/yy"))
End Sub

Private Sub Grd_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   Grd.Col = 8
   If Grd.row >= 1 Then
      If Grd.row = 0 Then Grd.row = 1
      If Grd.CellPicture = Chkyes.Picture Then
         Set Grd.CellPicture = Chkno.Picture
      Else
         Set Grd.CellPicture = Chkyes.Picture
      End If
   End If
End Sub


Private Sub optDateRange_Click()
   fraDateRange.Enabled = optDateRange.Value
   fraMO.Enabled = Not optDateRange.Value
End Sub

Private Sub optMO_Click()
   fraDateRange.Enabled = optDateRange.Value
   fraMO.Enabled = Not optDateRange.Value
End Sub

Private Function FindNewReCostValue(ByVal strMoPartRef As String, ByVal strMoRun As String, _
                           ByRef cUnitCost As Currency)
   'get the list of all completed or closed MO's in the date range
   Dim clsMO As New ClassMO
   clsMO.LoggingEnabled = False
   'Dim rdo As ADODB.Recordset
   Dim success As Boolean
   
   clsMO.DiagnoseOnly = 1
   clsMO.PartNumber = strMoPartRef
   clsMO.RunNumber = strMoRun
   success = clsMO.GetCalculatedMoCost(cUnitCost)
   
End Function

