VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form ShopSHf07 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Re-Schedule MOs"
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7275
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   7275
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkPC 
      Caption         =   "PC"
      Height          =   255
      Left            =   4080
      TabIndex        =   25
      Top             =   1440
      Width           =   615
   End
   Begin VB.CheckBox chkSelAllPrt 
      Caption         =   "Select All Part Number"
      Height          =   255
      Left            =   4920
      TabIndex        =   24
      Top             =   5160
      Width           =   2055
   End
   Begin VB.TextBox txtPtr 
      Height          =   285
      Left            =   3600
      TabIndex        =   23
      Top             =   5520
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.CommandButton cmdSel 
      Caption         =   "S&elect Part Number"
      Height          =   315
      Left            =   5160
      TabIndex        =   8
      ToolTipText     =   "Select Parts"
      Top             =   1800
      Width           =   1635
   End
   Begin VB.CheckBox chkPL 
      Caption         =   "PL"
      Height          =   255
      Left            =   3480
      TabIndex        =   7
      Top             =   1440
      Width           =   615
   End
   Begin VB.CheckBox chkPP 
      Caption         =   "PP"
      Height          =   255
      Left            =   2760
      TabIndex        =   6
      Top             =   1440
      Width           =   615
   End
   Begin VB.CheckBox chkRL 
      Caption         =   "RL"
      Height          =   255
      Left            =   2040
      TabIndex        =   5
      Top             =   1440
      Width           =   615
   End
   Begin VB.CheckBox chkSC 
      Caption         =   "SC"
      Height          =   255
      Left            =   1320
      TabIndex        =   4
      Top             =   1440
      Width           =   615
   End
   Begin VB.ComboBox cmbCls 
      Height          =   315
      Left            =   3600
      Sorted          =   -1  'True
      TabIndex        =   1
      Tag             =   "3"
      ToolTipText     =   "Select Product Class From List"
      Top             =   480
      Width           =   855
   End
   Begin VB.ComboBox cmbCde 
      Height          =   315
      Left            =   1320
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Select Product Code From List"
      Top             =   480
      Width           =   1215
   End
   Begin VB.ComboBox txtBeg 
      Height          =   315
      Left            =   1320
      TabIndex        =   2
      Tag             =   "4"
      Top             =   840
      Width           =   1215
   End
   Begin VB.ComboBox txtEnd 
      Height          =   315
      Left            =   3600
      TabIndex        =   3
      Tag             =   "4"
      Top             =   840
      Width           =   1215
   End
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   4200
      TabIndex        =   16
      ToolTipText     =   "Select Part Number"
      Top             =   5520
      Visible         =   0   'False
      Width           =   3180
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "ShopSHf07.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   14
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdSch 
      Caption         =   "&Schedule"
      Height          =   315
      Left            =   6240
      TabIndex        =   10
      TabStop         =   0   'False
      ToolTipText     =   "Update Entries and Re-Schedule"
      Top             =   720
      Width           =   875
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   6240
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   120
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   4440
      Top             =   5160
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   5670
      FormDesignWidth =   7275
   End
   Begin MSComctlLib.ProgressBar PrgBar 
      Height          =   300
      Left            =   240
      TabIndex        =   15
      Top             =   5160
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSFlexGridLib.MSFlexGrid grd 
      Height          =   2895
      Left            =   240
      TabIndex        =   9
      ToolTipText     =   "Click The Row To Select A Partnumber to Re-Schedule MO"
      Top             =   2160
      Width           =   6585
      _ExtentX        =   11615
      _ExtentY        =   5106
      _Version        =   393216
      Rows            =   10
      Cols            =   4
      FixedCols       =   0
      BackColorSel    =   -2147483640
      FocusRect       =   0
      HighLight       =   0
      GridLinesFixed  =   1
      ScrollBars      =   2
   End
   Begin VB.Image Chkno 
      Height          =   180
      Left            =   5160
      Picture         =   "ShopSHf07.frx":07AE
      Stretch         =   -1  'True
      Top             =   5160
      Visible         =   0   'False
      Width           =   350
   End
   Begin VB.Image Chkyes 
      Height          =   180
      Left            =   3720
      Picture         =   "ShopSHf07.frx":0805
      Stretch         =   -1  'True
      Top             =   5160
      Visible         =   0   'False
      Width           =   350
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Codes"
      Height          =   285
      Index           =   3
      Left            =   120
      TabIndex        =   22
      Top             =   480
      Width           =   1155
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Classes"
      Height          =   285
      Index           =   2
      Left            =   2640
      TabIndex        =   21
      Top             =   480
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Through"
      Height          =   255
      Index           =   5
      Left            =   2640
      TabIndex        =   20
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Start Date"
      Height          =   255
      Index           =   8
      Left            =   480
      TabIndex        =   19
      Top             =   840
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank for All)"
      Height          =   255
      Index           =   11
      Left            =   4920
      TabIndex        =   18
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank for All)"
      Height          =   255
      Index           =   1
      Left            =   4920
      TabIndex        =   17
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Run Status"
      Height          =   255
      Left            =   360
      TabIndex        =   13
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   11
      Top             =   1800
      Width           =   975
   End
End
Attribute VB_Name = "ShopSHf07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
Option Explicit
Dim bOnLoad As Byte

Private txtKeyPress(4) As New EsiKeyBd
Private txtGotFocus(4) As New EsiKeyBd
Private txtKeyDown(2) As New EsiKeyBd


Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub



Private Sub chkSelAllPrt_Click()
    
    Dim iList As Integer
    Grd.Col = 0
    For iList = 1 To Grd.Rows - 1
      ' Only if the part is checked
      Grd.row = iList
      Set Grd.CellPicture = Chkyes.Picture
    Next
End Sub

Private Sub cmbCde_LostFocus()
    If Trim(cmbCde) = "" Then cmbCde = "ALL"
End Sub

Private Sub cmbCls_LostFocus()
    If Trim(cmbCls) = "" Then cmbCls = "ALL"
End Sub

Private Sub cmbPrt_Click()
'   bGoodPart = GetPart(True)
   
End Sub

Private Sub cmbPrt_LostFocus()
'   cmbPrt = CheckLen(cmbPrt, 30)
'   If bCanceled Then Exit Sub
'   If Len(Trim(cmbPrt)) > 0 Then bGoodPart = GetPart(False)
   If Trim(cmbPrt) = "" Then cmbPrt = "ALL"
End Sub

Private Sub cmdCan_Click()
   Form_Deactivate
   
End Sub


Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 4171
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub cmdSch_Click()
   Dim bResponse As Byte
   Dim sMsg As String
   
   sMsg = "Reschedule MO Based On Latest Information?"
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      MouseCursor 13
      PrgBar.Visible = True
      'txtPtr.Visible = True
      BackSchedule
   End If
   PrgBar.Visible = False
   txtPtr.Visible = False
   MouseCursor 0
End Sub


Private Sub cmdSel_Click()
    chkSelAllPrt.Value = 0
    FillPartGrid
    MouseCursor 0
End Sub

Private Sub Form_Activate()
    Dim bGoodCal As Boolean
    
    Dim bGoodCoCal As Boolean
    
    If bOnLoad Then
    
      cmbCde.AddItem "ALL"
      FillProductCodes
      If Trim(cmbCde) = "" Then cmbCde = cmbCde.List(0)
      
      cmbCls.AddItem "ALL"
      FillProductClasses
      If Trim(cmbCls) = "" Then cmbCls = cmbCls.List(0)
        
      'FillPartGrid
      
      FillRuns Me, "NOT LIKE 'C%'"
      bGoodCal = GetCenterCalendar(Me)
      bGoodCoCal = GetCompanyCalendar()
      If bGoodCoCal = 0 Then
        MsgBox "There Is No Company Calendar For The Period.", _
            vbInformation, Caption
        CapaCPe04a.Show
        Unload Me
        Exit Sub
    End If
  End If
  bOnLoad = 0
  MouseCursor 0
   
End Sub

Private Sub txtBeg_DropDown()
   ShowCalendarEx Me
End Sub

Private Sub txtBeg_LostFocus()
   If Len(Trim(txtBeg)) = 0 Then txtBeg = "ALL"
   If txtBeg <> "ALL" Then txtBeg = CheckDateEx(txtBeg)
End Sub

Private Sub txtend_DropDown()
   ShowCalendarEx Me
End Sub

Private Sub txtEnd_LostFocus()
   If Len(Trim(txtEnd)) = 0 Then txtEnd = "ALL"
   If txtEnd <> "ALL" Then txtEnd = CheckDateEx(txtEnd)
End Sub

Private Sub GetMRPDates()
   Dim RdoDte As ADODB.Recordset
    sSql = "SELECT MIN(MRP_PARTDATERQD) FROM MrplTable WHERE " _
           & "MRP_TYPE>" & MRPTYPE_BeginningBalance
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoDte, ES_FORWARD)
   If bSqlRows Then
      With RdoDte
         If Not IsNull(.Fields(0)) Then
            txtBeg = Format(.Fields(0), "mm/dd/yy")
         Else
            txtBeg = Format(ES_SYSDATE, "mm/dd/yy")
         End If
         ClearResultSet RdoDte
      End With
   End If
   txtBeg.ToolTipText = "Earliest Date By Default"
   
   sSql = "SELECT MAX(MRP_PARTDATERQD) FROM MrplTable "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoDte, ES_FORWARD)
   If bSqlRows Then
      With RdoDte
         If Not IsNull(.Fields(0)) Then
            txtEnd = Format(.Fields(0), "mm/dd/yy")
         Else
            txtEnd = Format(ES_SYSDATE, "mm/dd/yy")
         End If
         ClearResultSet RdoDte
      End With
   End If
   txtEnd.ToolTipText = "Latest Date By Default"
   
End Sub
Private Sub Form_Deactivate()
   Unload Me
   
End Sub


Private Sub Form_Load()
   FormLoad Me
   bOnLoad = 1
   FormatControls
   txtEnd = "ALL"
   txtBeg = "ALL"
    
      With Grd
      .ColAlignment(0) = 0
      .ColAlignment(1) = 0
      .Rows = 1
      .row = 0
      .Col = 0
      .Text = "Sel"
      .Col = 1
      .Text = "Part Number"
      .Col = 2
      .Text = "Description"
      '.Col = 3
      '.Text = "Run Status"
      .ColWidth(0) = 350
      .ColWidth(1) = 2300
      .ColWidth(2) = 3600
      '.ColWidth(3) = 600
      
   End With
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   Set ShopSHf07 = Nothing
   
End Sub

Private Sub FillPartGrid()
    Dim RdoGrd As ADODB.Recordset
    Dim strRunStat As String
    Dim strCls As String
    Dim strProdCode As String
    Dim strSC As String
    Dim strRL As String
    Dim strPL As String
    Dim strPP As String
    Dim strPC As String
    Dim strBegDt As String
    Dim strEndDt As String
    
    On Error Resume Next
    
    Grd.Rows = 1
    On Error GoTo DiaErr1
    
    If Trim(cmbCde) = "ALL" Then
        strProdCode = "%"
    Else
        strProdCode = Trim(cmbCde)
    End If
    
    If Trim(cmbCls) = "ALL" Then
        strCls = "%"
    Else
        strCls = Trim(cmbCls)
    End If
    
    If (chkSC.Value = 1) Then
        strSC = "SC"
    Else
        strSC = ""
    End If
    
    If (chkRL.Value = 1) Then
        strRL = "RL"
    Else
        strRL = ""
    End If
    
    If (chkPL.Value = 1) Then
        strPL = "PL"
    Else
        strPL = ""
    End If
    
    If (chkPP.Value = 1) Then
        strPP = "PP"
    Else
        strPP = ""
    End If
    
    If (chkPC.Value = 1) Then
        strPC = "PC"
    Else
        strPC = ""
    End If
    
    ' IF the begin date and end date are ALL
    ' Get the max and min dates from MRPL Table
    
    If (txtBeg = "ALL") And (txtEnd = "ALL") Then
        GetMRPDates
    End If
    
    Dim sSql1 As String
    Dim sSql2  As String
    
   sSql1 = "SELECT DISTINCT PARTNUM,PADESC FROM PartTable, RunsTable,MrplTable  " & _
             "WHERE PARTREF=RUNREF AND PARTREF = MRP_PARTREF AND RUNSTATUS NOT LIKE 'C_'" & _
            "AND RUNSTATUS IN ('" & strSC & "','" & strRL & "','" & strPL & "','" & strPP & "','" & strPC & "')" & _
            " AND PAPRODCODE LIKE '" & strProdCode & "'"

    sSql2 = " AND PACLASS LIKE '" & strCls & "'" & _
            " AND MRP_PARTDATERQD BETWEEN '" & txtBeg & "' AND '" & txtEnd & "'" & _
            " ORDER BY PARTNUM"
            
    sSql = sSql1 & sSql2
   
   Debug.Print sSql
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoGrd, ES_FORWARD)
   If bSqlRows Then
      With RdoGrd
         Do Until .EOF
            Grd.Rows = Grd.Rows + 1
            Grd.row = Grd.Rows - 1
            Grd.Col = 0
            Set Grd.CellPicture = Chkno.Picture
            
            Grd.Col = 1
            Grd.Text = "" & Trim(!PartNum)
            Grd.Col = 2
            Grd.Text = "" & Trim(!PADESC)
            'grd.Col = 3
            'grd.Text = "" & Trim(!RUNSTATUS)
            ' 9/9/2013
            ' MM If Grd.Rows > 300 Then Exit Do
            .MoveNext
         Loop
         ClearResultSet RdoGrd
      End With
   Else
      MsgBox "There Are No Parts for this criteria.", _
         vbInformation, Caption
   End If
   Set RdoGrd = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "FillPartGrid"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub grd_KeyPress(KeyAscii As Integer)
   On Error Resume Next
   If KeyAscii = vbKeyReturn Or KeyAscii = vbKeySpace Then
      Grd.Col = 0
      If Grd.CellPicture = Chkyes.Picture Then
         Set Grd.CellPicture = Chkno.Picture
      Else
         Set Grd.CellPicture = Chkyes.Picture
      End If
   End If
   
End Sub

Private Sub Grd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   On Error Resume Next
   Grd.Col = 0
   If Grd.CellPicture = Chkyes.Picture Then
      Set Grd.CellPicture = Chkno.Picture
   Else
      Set Grd.CellPicture = Chkyes.Picture
   End If
   
End Sub

Private Function BackSchedule()

    'MouseCursor ccHourglass
    Dim RdoMon As ADODB.Recordset
    Dim strPassedMo As String
    Dim strPartNumber As String
    Dim dtSched As Date
    Dim lRunno As Long
    Dim lQty  As Long
    Dim lRowcnt As Long
    Dim lrows  As Long
    Dim lprg As Long
    Dim iList As Integer
    
    Dim strSC As String
    Dim strRL As String
    Dim strPL As String
    Dim strPP As String
    Dim strPC As String
    
    On Error GoTo DiaErr1
    'MouseCursor 13
   
    If (chkSC.Value = 1) Then
        strSC = "SC"
    Else
        strSC = ""
    End If
    
    If (chkRL.Value = 1) Then
        strRL = "RL"
    Else
        strRL = ""
    End If
    
    If (chkPL.Value = 1) Then
        strPL = "PL"
    Else
        strPL = ""
    End If
    
    If (chkPP.Value = 1) Then
        strPP = "PP"
    Else
        strPP = ""
    End If
   
    If (chkPC.Value = 1) Then
        strPC = "PC"
    Else
        strPC = ""
    End If
   
   
    ' Go throught all the record int he grid and re-schedule MO
    For iList = 1 To Grd.Rows - 1
      Grd.Col = 0
      Grd.row = iList
      ' Only if the part is checked
      If Grd.CellPicture = Chkyes.Picture Then
        Grd.Col = 1
        strPartNumber = Compress(Grd.Text)
        
        sSql = "SELECT PARTREF,PARTNUM,PADESC,PAUNITS,PALEVEL,PARUN,RUNREF," _
            & "RUNNO,RUNSTATUS,RUNQTY,RUNSCHED,RUNPRIORITY,RUNDIVISION " _
            & "FROM PartTable,RunsTable WHERE PARTREF  = '" & strPartNumber & "' " _
            & "AND PARTREF=RUNREF AND RUNSTATUS NOT LIKE 'C%' " _
            & "AND RUNSTATUS IN ('" & strSC & "','" & strRL & "','" & strPL & "','" & strPP & "','" & strPC & "')"
      
      Debug.Print sSql
      
      
        bSqlRows = clsADOCon.GetDataSet(sSql, RdoMon, ES_STATIC)
        lRowcnt = 0
        PrgBar.Value = 5
        
        If bSqlRows Then
          With RdoMon
            lrows = RdoMon.RecordCount
            Do Until .EOF
                Dim mo As New ClassMO
                ' Get the run number
                txtPtr = "" & Trim(!PartNum)
                strPassedMo = Compress(Trim(!PartNum))
                lQty = Trim(!RUNQTY)
                dtSched = "" & Format(!RUNSCHED, "mm/dd/yy")
                
                lRunno = Trim(!Runno)
                mo.ScheduleOperations strPassedMo, lRunno, CCur(lQty), dtSched, False
                Set mo = Nothing
                lRowcnt = lRowcnt + 1
                If (lrows > 0) Then
                    lprg = CInt((lRowcnt * 100) / lrows)
                    If (lprg > 1) Then
                        PrgBar.Value = CInt(lprg)
                    End If
                End If
               .MoveNext
            Loop
            ClearResultSet RdoMon
          End With
        End If
        Set RdoMon = Nothing
   
      End If
    Next

   MouseCursor 0
   Exit Function

DiaErr1:
   sProcName = "BackSchedule"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function


'Private Function BackSchedule()
'
'    MouseCursor ccHourglass
'    Dim RdoMon As ADODB.Recordset
'    Dim sPassedMo As String
'    Dim sPartNumber As String
'    Dim dtSched As Date
'    Dim lRunno As Long
'    Dim lQty  As Long
'    Dim lRowcnt As Long
'    Dim lRows  As Long
'    Dim lprg As Long
'
'    sPartNumber = Compress(cmbPrt)
'
'    On Error GoTo DiaErr1
'    MouseCursor 13
'
'    If Trim(cmbPrt) = "ALL" Then
'
'        sSql = "SELECT PARTREF,PARTNUM,PADESC,PAUNITS,PALEVEL,PARUN,RUNREF," _
'            & "RUNNO,RUNSTATUS,RUNQTY,RUNSCHED,RUNPRIORITY,RUNDIVISION " _
'            & "FROM PartTable,RunsTable WHERE PARTREF LIKE '%'" _
'            & "AND PARTREF=RUNREF AND RUNSTATUS NOT LIKE 'C%'"
'    Else
'        sSql = "SELECT PARTREF,PARTNUM,PADESC,PAUNITS,PALEVEL,PARUN,RUNREF," _
'            & "RUNNO,RUNSTATUS,RUNQTY,RUNSCHED,RUNPRIORITY,RUNDIVISION " _
'            & "FROM PartTable,RunsTable WHERE PARTREF  = '" & sPartNumber & "'" _
'            & "AND PARTREF=RUNREF AND RUNSTATUS NOT LIKE 'C%'"
'    End If
'
'
'    bsqlrows = clsadocon.getdataset(ssql, RdoMon, ES_STATIC)
'    lRowcnt = 0
'    PrgBar.value = 5
'
'    If bSqlRows Then
'      With RdoMon
'        lRows = RdoMon.RowCount
'        Do Until .EOF
'            Dim mo As New ClassMO
'            ' Get the run number
'            cmbPrt = "" & Trim(!PartNum)
'            sPassedMo = Compress(Trim(!PartNum))
'            lQty = Trim(!RUNQTY)
'            dtSched = "" & Format(!RUNSCHED, "mm/dd/yy")
'
'            lRunno = Trim(!Runno)
'            mo.ScheduleOperations sPassedMo, lRunno, CCur(lQty), dtSched, False
'            Set mo = Nothing
'            lRowcnt = lRowcnt + 1
'            If (lRows > 0) Then
'                lprg = CInt((lRowcnt * 100) / lRows)
'                If (lprg > 1) Then
'                    PrgBar.value = CInt(lprg)
'                End If
'            End If
'           .MoveNext
'        Loop
'        ClearResultSet RdoMon
'      End With
'   End If
'   Set RdoMon = Nothing
'
'   MouseCursor 0
'   Exit Function
'
'DiaErr1:
'   sProcName = "BackSchedule"
'   CurrError.Number = Err.Number
'   CurrError.Description = Err.Description
'   DoModuleErrors Me
'
'End Function

