VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form EstiEse06a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Enter Engineering Time Charges"
   ClientHeight    =   7860
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10620
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7860
   ScaleWidth      =   10620
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdLogin 
      Caption         =   "Log In"
      Height          =   360
      Left            =   2640
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   180
      Width           =   1065
   End
   Begin VB.TextBox txtComments 
      Height          =   1155
      Left            =   1440
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   2580
      Width           =   6255
   End
   Begin VB.CommandButton cmdAdd 
      Cancel          =   -1  'True
      Caption         =   "Add"
      Height          =   360
      Left            =   3720
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   3840
      Width           =   1065
   End
   Begin VB.TextBox txtHours 
      Height          =   285
      Left            =   1440
      TabIndex        =   5
      Tag             =   "2"
      ToolTipText     =   "Tool Description (40 chars)"
      Top             =   2160
      Width           =   585
   End
   Begin VB.TextBox txtPIN 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1440
      PasswordChar    =   "*"
      TabIndex        =   0
      Tag             =   "2"
      ToolTipText     =   "Tool Description (40 chars)"
      Top             =   180
      Width           =   1065
   End
   Begin VB.ComboBox cmbDate 
      Height          =   315
      Left            =   1440
      TabIndex        =   2
      Tag             =   "4"
      Top             =   960
      Width           =   1335
   End
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   1440
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Tag             =   "3"
      ToolTipText     =   "Contains Part Numbers With Manufacturing Orders"
      Top             =   1440
      Width           =   3540
   End
   Begin VB.ComboBox cmbRun 
      Height          =   315
      Left            =   5400
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Tag             =   "3"
      ToolTipText     =   "Contains Part Numbers With Manufacturing Orders"
      Top             =   1440
      Width           =   1140
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Close"
      Height          =   360
      Left            =   6420
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   3255
      Left            =   300
      TabIndex        =   22
      Top             =   4440
      Width           =   10035
      _ExtentX        =   17701
      _ExtentY        =   5741
      _Version        =   393216
      Rows            =   0
      FixedRows       =   0
      RowHeightMin    =   285
      BackColorBkg    =   -2147483633
      WordWrap        =   -1  'True
      AllowBigSelection=   0   'False
   End
   Begin VB.Label lblOperation 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   6900
      TabIndex        =   21
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Operation"
      Height          =   255
      Index           =   6
      Left            =   5040
      TabIndex        =   20
      Top             =   1860
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Comments"
      Height          =   285
      Index           =   3
      Left            =   240
      TabIndex        =   19
      Top             =   2580
      Width           =   795
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Recent Charges"
      Height          =   285
      Index           =   7
      Left            =   300
      TabIndex        =   18
      Top             =   4080
      Width           =   1275
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Hours (X.XX)"
      Height          =   285
      Index           =   5
      Left            =   240
      TabIndex        =   17
      Top             =   2160
      Width           =   1035
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Engineer"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   16
      Top             =   600
      Width           =   855
   End
   Begin VB.Label lblEngineer 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "<Not logged in>"
      Height          =   285
      Left            =   1440
      TabIndex        =   15
      Top             =   600
      Width           =   3255
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "PIN"
      Height          =   285
      Index           =   28
      Left            =   240
      TabIndex        =   14
      Top             =   180
      Width           =   795
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   13
      Top             =   960
      Width           =   855
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1440
      TabIndex        =   12
      Top             =   1800
      Width           =   3255
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Run"
      Height          =   255
      Index           =   1
      Left            =   5040
      TabIndex        =   11
      Top             =   1500
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   10
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   6900
      TabIndex        =   9
      Top             =   1440
      Width           =   495
   End
End
Attribute VB_Name = "EstiEse06a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

    Enum GRID_COL
        COL_DELETE
        COL_DATE
        COL_HOURS
        COL_PARTNUM
        COL_PARTDESC
        COL_RUN
        COL_OP
        COL_ENTERED
    End Enum


Dim rdoQry As ADODB.Command
Dim AdoParameter1 As ADODB.Parameter

'Dim DbDoc   As Recordset 'Jet
'Dim DbPls   As Recordset 'Jet
Dim bPrinting As Boolean

Dim bGoodPart As Byte
Dim bGoodMO As Byte
Dim bOnLoad As Byte
Dim bTablesCreated As Byte
Dim bUserTypedRun As Byte

Dim sBomRev As String
Dim sRunPkstart As String
Dim sPartNumber As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub GetOptions()
   txtPIN = GetSetting("Esi2000", "EsiEngr", "EstiEse06a", "")
   If txtPIN <> "" Then
      GetEngineer
   End If
   
End Sub


Private Sub SaveOptions()
   Dim sOptions As String
   
   sOptions = txtPIN.Text
   SaveSetting "Esi2000", "EsiEngr", "EstiEse06a", Trim(sOptions)
'   SaveSetting "Esi2000", "EsiProd", "sh01all", Trim(chkSoAlloc.Value)
'   SaveSetting "Esi2000", "EsiProd", "sh01Printer", lblPrinter
   
End Sub

Private Sub cmbDate_DropDown()
   ShowCalendar Me
End Sub


Private Sub cmbPrt_Click()
   GetPart
End Sub

Private Sub cmbRun_Click()
   GetRun
End Sub


Private Sub cmbRun_KeyDown(KeyCode As Integer, Shift As Integer)
    bUserTypedRun = 1
End Sub

Private Sub cmdAdd_Click()

   Dim hours As Single
   If IsNumeric(txtHours.Text) Then
      hours = CSng(txtHours.Text)
   End If
   
   If hours <= 0 Or cmbPrt.Text = "" Or cmbRun.Text = "" _
      Or lblOperation = "" Or Not IsNumeric(txtHours.Text) Then
      
      MsgBox "Information missing or invalid.  Cannot proceed."
      Exit Sub
   End If
   
   'make sure there is an open time journal
   Dim rs As ADODB.Recordset
   Dim JournalID As String
   sSql = "select dbo.fnGetOpenJournalID('TJ', '" & cmbDate.Text & "')"
   bSqlRows = clsADOCon.GetDataSet(sSql, rs)
   If bSqlRows Then
      JournalID = rs(0)
      rs.Close
      Set rs = Nothing
   Else
      MsgBox "No open time journal for this date"
      rs.Close
      Set rs = Nothing
      Exit Sub
   End If
   
   cmdAdd.enabled = False
   If AddTimeCharge(JournalID) Then
      ClearControls
      RefreshGrid
   End If

   cmdAdd.enabled = True
   
End Sub

Private Function AddTimeCharge(JournalID As String) As Boolean
   AddTimeCharge = False

   sSql = "exec AddEngrTimeCharge " & vbCrLf _
      & txtPIN.Text & "," & vbCrLf _
      & "'" & cmbDate.Text & "'," & vbCrLf _
      & "'" & cmbPrt.Text & "'," & vbCrLf _
      & cmbRun.Text & "," & vbCrLf _
      & lblOperation.Caption & "," & vbCrLf _
      & txtHours.Text & "," & vbCrLf _
      & "'" & Replace(Trim(txtComments.Text), "'", "''") & "'," & vbCrLf _
      & "'" & JournalID & "'"
      
   Dim success As Boolean
   success = clsADOCon.ExecuteSql(sSql)
    
   If success Then
      AddTimeCharge = True
   Else
      MsgBox "Cannot add time charge.", _
         vbInformation, Caption
      Err.Clear
      Set clsADOCon = Nothing
      Exit Function
   End If
   
End Function
   
'   sSql = "exec create procedure AddEngrTimeCharge " & vbCrLf _
'      & txpin & "," & vbCrLf _
'      & "'" & cmbDate & "'," & vbCrLf _
'      & "'" & cmbPrt.Text & "'," & vbCrLf _
'      & cmbRun.Text & "," & vbCrLf _
'      & lblOperation.Text & vbCrLf _
'      & txtHours.Text & "," & vbCrLf _
'      & "'" & txtcomment.Text & "'"
'
'   Dim success As Boolean
'   success = clsADOCon.ExecuteSql(sSql)
'
'   If Not success Then
'      MsgBox "Cannot add time charge.  An open time journal is required.", _
'         vbInformation, Caption
'      Err.Clear
'      Set clsADOCon = Nothing
'      Exit Sub
'   End If

'Private Sub cmbRun_LostFocus()
'   cmbRun = CheckLen(cmbRun, 5)
'   If Val(cmbRun) > 32767 Then cmbRun = "32767"
'   cmbRun = Format(Abs(Val(cmbRun)), "####0")
'   GetThisRun
'
'End Sub
'

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdLogin_Click()
   If cmdLogin.Caption = "Log In" Then
      GetEngineer
   Else
      Logout
      EnableControls False
      SaveOptions
   End If
End Sub

Private Sub Logout()
   txtPIN = ""
   lblEngineer = "<Not Logged In>"
   cmbDate.Text = ""
   ClearControls
End Sub

Private Sub ClearControls()
   cmbPrt.ListIndex = -1
   cmbRun.Clear
   txtHours.Text = ""
   txtComments.Text = ""
   grid.Clear
   grid.Rows = 0
End Sub

Private Sub Form_Activate()
   If bOnLoad Then
      bOnLoad = 0
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   EnableControls False
   FormLoad Me
   txtComments = ""
   txtHours = ""
   txtPIN = ""
   
   'FormatControls
   bUserTypedRun = 0
   
   GetOptions
   bTablesCreated = 0
   sSql = "SELECT DISTINCT PARTREF,PARTNUM" & vbCrLf _
          & "FROM PartTable parts" & vbCrLf _
          & "JOIN RunsTable runs on runs.RUNREF = parts.PARTREF" & vbCrLf _
          & "WHERE RUNSTATUS NOT IN ('CL','CA','CO')"
   LoadComboBox cmbPrt, -1, False
   bOnLoad = 1
   
End Sub




Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveOptions
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   Set AdoParameter1 = Nothing
   Set rdoQry = Nothing
   FormUnload
   Set EstiEse06a = Nothing
   
End Sub




'Private Function GetRuns() As Byte
'   Dim RdoRns As ADODB.Recordset
'   Dim iOriginalRun As Integer
'   Dim bOriginalRunFound As Byte
'
'   bOriginalRunFound = 0
'
'   On Error GoTo DiaErr1
'   iOriginalRun = Val(cmbRun)
'   MouseCursor 13
'   cmbRun.Clear
'   sPartNumber = Compress(cmbPrt)
'   rdoQry.Parameters(0).Value = sPartNumber
''   rdoQry(0) = sPartNumber
'   bSqlRows = clsADOCon.GetQuerySet(RdoRns, rdoQry)
'   If bSqlRows Then
'      With RdoRns
'         cmbRun = Format(!Runno, "####0")
'         lblDsc = "" & Trim(!PADESC)
'         lblTyp = Format(!PALEVEL, "#")
'         Do Until .EOF
'            AddComboStr cmbRun.hwnd, Format$(!Runno, "####0")
'            If iOriginalRun = !Runno Then bOriginalRunFound = 1
'            .MoveNext
'         Loop
'         ClearResultSet RdoRns
'      End With
'      cmbRun = cmbRun.List(cmbRun.ListCount - 1)
'      GetRuns = True
'      GetThisRun
'   Else
'      sPartNumber = ""
'      GetRuns = False
'   End If
'   MouseCursor 0
'   Set RdoRns = Nothing
'   Exit Function
'
'DiaErr1:
'   sProcName = "getruns"
'   CurrError.Number = Err.Number
'   CurrError.Description = Err.Description
'   DoModuleErrors Me
'
'End Function
'
'Private Sub lblQty_Click()
'   'run qty
'
'End Sub


'Private Sub GetThisRun()
'   Dim RdoRun As ADODB.Recordset
'
'   On Error GoTo DiaErr1
'   sSql = "SELECT RUNREF,RUNNO,RUNSTATUS,RUNPKSTART,RUNQTY FROM RunsTable WHERE " _
'          & "RUNREF='" & Compress(cmbPrt) & "' AND " _
'          & "RUNNO=" & cmbRun & " "
'   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRun, ES_FORWARD)
'   If bSqlRows Then
'      With RdoRun
'         lblSta = "" & Trim(!RUNSTATUS)
'         If Not IsNull(!RUNPKSTART) Then
'            sRunPkstart = Format(!RUNPKSTART, "mm/dd/yy")
'         Else
'            sRunPkstart = Format(ES_SYSDATE, "mm/dd/yy")
'         End If
'         ClearResultSet RdoRun
'      End With
'   End If
'   Set RdoRun = Nothing
'   Exit Sub
'
'DiaErr1:
'   sProcName = "getthisrun"
'   CurrError.Number = Err.Number
'   CurrError.Description = Err.Description
'   DoModuleErrors Me
'
'End Sub
'
Public Sub EnableControls(enable As Boolean)
   cmbDate.enabled = enable
   cmbPrt.enabled = enable
   cmbRun.enabled = enable
   txtHours.enabled = enable
   txtComments.enabled = enable
   cmdAdd.enabled = enable
   txtPIN.enabled = Not enable
   If enable Then cmdLogin.Caption = "Log Out" Else cmdLogin.Caption = "Log In"
End Sub

Public Function GetEngineer() As Boolean
   
   If Not IsNumeric(txtPIN.Text) Then Exit Function
   Dim prdoEmpl As ADODB.Recordset, logins As String
   sSql = "SELECT PREMLSTNAME,PREMFSTNAME" & vbCrLf _
          & "FROM EmplTable " & vbCrLf _
          & "WHERE PREMNUMBER = " & txtPIN.Text & vbCrLf _
          & "AND PREMENGINEER = 1"
          
   bSqlRows = clsADOCon.GetDataSet(sSql, prdoEmpl)
   If bSqlRows Then
      With prdoEmpl
         lblEngineer = Trim(!PREMFSTNAME) & " " & Trim(!PREMLSTNAME)
      End With
      EnableControls True
      GetEngineer = True
      SaveOptions
      cmbDate.Text = Format(Now, "MM/dd/yy")
      RefreshGrid
   Else
      If Trim(txtPIN.Text) = "" Then
         lblEngineer.Caption = "<Not logged in>"
      Else
         lblEngineer.Caption = "<Not a valid engineer PIN>"
      End If
      txtPIN = ""
      SaveOptions
      GetEngineer = False
   End If
End Function

Private Sub GetPart()
   lblDsc = ""
   lblStatus = ""
   lblOperation = ""
   cmbRun.Clear
   
   Dim rs As ADODB.Recordset
   sSql = "select RTRIM(PADESC) as PADESC from PartTable where PARTREF = '" & Compress(cmbPrt.Text) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, rs, ES_FORWARD)
   If bSqlRows Then
      With rs
         lblDsc = !PADESC
      End With
   End If
   
   sSql = "select RUNNO from RunsTable" & vbCrLf _
      & "where RUNREF = '" & Compress(cmbPrt.Text) & "'" & vbCrLf _
      & "and RUNSTATUS NOT IN ('CL','CA','CO')" & vbCrLf _
      & "order by RUNNO"
   LoadComboBox cmbRun, -1

End Sub

Private Sub GetRun()
   lblStatus = ""
   lblOperation = ""
   Dim rs As ADODB.Recordset
   sSql = "select top 1 RUNSTATUS, op.OPNO from RunsTable run" & vbCrLf _
      & "left join RnopTable op on op.OPREF = run.RUNREF and op.OPRUN = run.RUNNO" & vbCrLf _
      & "where RUNREF = '" & cmbPrt.Text & "' and RUNNO = " & cmbRun.Text & vbCrLf _
      & "order by op.OPNO"
   bSqlRows = clsADOCon.GetDataSet(sSql, rs)
   If bSqlRows Then
      With rs
         lblStatus = !RUNSTATUS
         lblOperation = !OPNO
      End With
   End If
End Sub

Private Sub grid_Click()
   Dim hours As Single
   Dim dt As String
   Dim part As String
   Dim run As Integer
   Dim op As Integer
   Dim entered As Date
   Dim row As Integer
   Dim col As Integer
   
   row = grid.MouseRow     'these numbers aren't accurate when a breakpoint is hit
   col = grid.MouseCol
   
   If row = 0 Then Exit Sub
   If col = COL_DELETE Then
      dt = grid.TextMatrix(row, COL_DATE)
      hours = CSng(grid.TextMatrix(row, COL_HOURS))
      part = grid.TextMatrix(row, COL_PARTNUM)
      run = CInt(grid.TextMatrix(row, COL_RUN))
      entered = CDate(grid.TextMatrix(row, COL_ENTERED))
      Dim msg As String
      msg = "Delete " & dt & " " & CStr(hours) & " hour time charge for " & part & " run " & CStr(run)
      If MsgBox(msg, vbYesNo) = vbYes Then
         sSql = "delete from TcitTable where TCEMP = " & txtPIN.Text & vbCrLf _
            & "and TCHOURS = " & CStr(hours) & vbCrLf _
            & "and TCSOURCE = 'Engr'" & vbCrLf _
            & "and TCPARTREF = '" & Compress(part) & "'" & vbCrLf _
            & "and TCRUNNO = " & CStr(run) & vbCrLf _
            & "and TCENTERED = '" & CStr(entered) & "'"
            
         Dim success As Boolean
         success = clsADOCon.ExecuteSql(sSql)
          
         If Not success Then
            MsgBox "Unable to delete time charge.", _
               vbInformation, Caption
            Err.Clear
            Set clsADOCon = Nothing
            Exit Sub
         Else
            RefreshGrid
         End If
      
      End If
   End If
   'part = grid.TextMatrix(grid.Row, grid.col)
End Sub

'Public Function IsValidTime(time As Variant)
'   On Error Resume Next
'   Dim n As Long
'   n = DatePart("n", time)
'   If Err Then
'      IsValidTime = False
'   Else
'      IsValidTime = True
'   End If
'End Function

Private Sub txtHours_LostFocus()
   
   Dim c
   c = CheckDecimal(txtHours.Text, "#0.00", True)
   If c = "*" Then
      txtHours.SetFocus
   Else
      txtHours = c
   End If

End Sub

Private Sub RefreshGrid()
    Dim Key As String
    'Key = cboLotID.Text
    grid.Clear
    grid.Rows = 0
   
    Dim rdo As ADODB.Recordset

   sSql = "SELECT TOP 10 'DELETE' as [Delete?], " & vbCrLf _
      & "CONVERT(varchar(10),TCSTARTTIME,101) AS [Date], " & vbCrLf _
      & "TCHOURS as Hrs, rtrim(PARTNUM) as [Part #], PADESC as Description," & vbCrLf _
      & "TCRUNNO as Run, TCOPNO as Op, TCENTERED as Entered" & vbCrLf _
      & "FROM TcitTable tc" & vbCrLf _
      & "JOIN PartTable pt on pt.PARTREF = tc.TCPARTREF" & vbCrLf _
      & "WHERE TCEMP = " & txtPIN & " and TCSOURCE = 'Engr'" & vbCrLf _
      & "ORDER BY TCSTARTTIME DESC, TCENTERED DESC" & vbCrLf
    bSqlRows = clsADOCon.GetDataSet(sSql, rdo, adUseClient)
    If bSqlRows Then
        Dim iCol As Integer
        Dim iRow As Integer
        Dim fld As ADODB.Field
        
        'insert headers
        With grid
            .cols = rdo.Fields.Count
            .Rows = rdo.RecordCount + 1
            .ColWidth(COL_DELETE) = 800
            .ColWidth(COL_DATE) = 1000
            .ColWidth(COL_HOURS) = 500
            .ColWidth(COL_PARTNUM) = 1800
            .ColAlignment(COL_PARTNUM) = 1 'LEFT center
            .ColWidth(COL_PARTDESC) = 3000
            .ColAlignment(COL_PARTDESC) = 1 'LEFT center
            .ColWidth(COL_RUN) = 500
            .ColWidth(COL_OP) = 500
            .ColWidth(COL_ENTERED) = 0    'don't show
            
            .FixedRows = 1
            
            Dim gridWidth As Integer
            'gridWidth = 360 'allow for scrollbar
            gridWidth = 120
            Dim i As Integer
            For i = 0 To .cols - 1
                gridWidth = gridWidth + .ColWidth(i)
            Next i
            .Width = gridWidth
            Dim minWidth As Integer
            minWidth = Me.txtComments.Left + txtComments.Width + 600
            If 2 * grid.Left + gridWidth > minWidth Then Me.Width = 2 * grid.Left + gridWidth Else Me.Width = minWidth

            .FixedRows = 1
            
           
            iCol = 0
            For Each fld In rdo.Fields
                 grid.TextMatrix(0, iCol) = fld.Name
                iCol = iCol + 1
            Next fld
        
            iRow = 1
            Do Until rdo.EOF
                iCol = 0
                For Each fld In rdo.Fields
                        If IsNull(fld.Value) = True Then
                            grid.TextMatrix(iRow, iCol) = vbNullString
                        Else
                            grid.TextMatrix(iRow, iCol) = fld.Value
                        End If
                    iCol = iCol + 1
                Next fld
                iRow = iRow + 1
                rdo.MoveNext
            Loop
            
        End With
    End If
End Sub

