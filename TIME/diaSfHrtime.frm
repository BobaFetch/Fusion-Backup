VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form diaSfHrtime 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Apply Shift Code to  Daily Time Charges"
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12630
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "diaSfHrtime.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   12630
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtAdjMin 
      Enabled         =   0   'False
      Height          =   315
      Left            =   4920
      TabIndex        =   28
      Top             =   1800
      Width           =   615
   End
   Begin VB.CheckBox chkAddRate 
      Height          =   255
      Left            =   360
      TabIndex        =   24
      Top             =   2400
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtSfBeg 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1320
      TabIndex        =   19
      Text            =   " :"
      Top             =   1320
      Width           =   735
   End
   Begin VB.TextBox txtSfEnd 
      Enabled         =   0   'False
      Height          =   315
      Left            =   3720
      TabIndex        =   18
      Text            =   " :"
      Top             =   1320
      Width           =   735
   End
   Begin VB.TextBox txtTotHrs 
      Enabled         =   0   'False
      Height          =   315
      Left            =   4920
      TabIndex        =   17
      Top             =   1320
      Width           =   615
   End
   Begin VB.TextBox txtSLEnd 
      Enabled         =   0   'False
      Height          =   315
      Left            =   3720
      TabIndex        =   16
      Text            =   " :"
      Top             =   1800
      Width           =   735
   End
   Begin VB.TextBox txtSLBeg 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1320
      TabIndex        =   15
      Text            =   " :"
      Top             =   1800
      Width           =   735
   End
   Begin VB.ComboBox cmbSfCd 
      Height          =   315
      Left            =   4200
      TabIndex        =   2
      Tag             =   "2"
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      Enabled         =   0   'False
      Height          =   315
      Left            =   8940
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Enter Updated Time Card"
      Top             =   1920
      Width           =   875
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   "C&lear"
      Enabled         =   0   'False
      Height          =   315
      Left            =   8040
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Cancel Time Card Entry"
      Top             =   1920
      Width           =   875
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   8760
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   240
      Width           =   875
   End
   Begin VB.ComboBox cmbEmp 
      Height          =   315
      Left            =   960
      TabIndex        =   1
      Tag             =   "1"
      ToolTipText     =   "Select From List Or Enter Number"
      Top             =   190
      Width           =   1095
   End
   Begin VB.ComboBox txtDte 
      Height          =   315
      Left            =   6360
      TabIndex        =   3
      Tag             =   "3"
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   7920
      Picture         =   "diaSfHrtime.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   6120
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   250
   End
   Begin MSFlexGridLib.MSFlexGrid Grd 
      Height          =   4455
      Left            =   240
      TabIndex        =   7
      ToolTipText     =   "Click To Select Or Scroll And Press Enter (Also Page Up And Page Down)"
      Top             =   2880
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   7858
      _Version        =   393216
      Rows            =   3
      Cols            =   11
      FixedRows       =   2
      FixedCols       =   0
      FocusRect       =   2
      ScrollBars      =   2
      MergeCells      =   1
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "* Employee Shift code applied"
      Height          =   255
      Index           =   11
      Left            =   360
      TabIndex        =   30
      Top             =   7440
      Width           =   1695
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Time Adjustment in minutes"
      Height          =   375
      Index           =   10
      Left            =   5640
      TabIndex        =   29
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   """TOT Hrs"" doesn't include Lunch"
      Height          =   255
      Index           =   9
      Left            =   4200
      TabIndex        =   27
      ToolTipText     =   "Additional rate"
      Top             =   2520
      Width           =   2535
   End
   Begin VB.Label lblAddRate 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2400
      TabIndex        =   26
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Additional Rate"
      Height          =   255
      Index           =   8
      Left            =   1200
      TabIndex        =   25
      ToolTipText     =   "Additional rate"
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Start Time"
      Height          =   255
      Index           =   5
      Left            =   480
      TabIndex        =   23
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "End Time"
      Height          =   255
      Index           =   4
      Left            =   2880
      TabIndex        =   22
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "LunchEnd Time"
      Height          =   255
      Index           =   6
      Left            =   2400
      TabIndex        =   21
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "LunchStart Time"
      Height          =   255
      Index           =   7
      Left            =   0
      TabIndex        =   20
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Shift Code"
      Height          =   255
      Index           =   3
      Left            =   3360
      TabIndex        =   14
      Top             =   240
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Employee"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   13
      Top             =   195
      Width           =   735
   End
   Begin VB.Label lblNme 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   960
      TabIndex        =   12
      Top             =   645
      Width           =   2415
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   11
      Top             =   645
      Width           =   615
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Card Date"
      Height          =   255
      Index           =   1
      Left            =   5520
      TabIndex        =   10
      Top             =   240
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Week Ending"
      Height          =   255
      Index           =   12
      Left            =   7680
      TabIndex        =   9
      Top             =   1245
      Width           =   1215
   End
   Begin VB.Label lblWen 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   8880
      TabIndex        =   8
      ToolTipText     =   "Week Ending (System Administration Setup)"
      Top             =   1245
      Width           =   855
   End
   Begin VB.Image Chkyes 
      Height          =   210
      Left            =   7800
      Picture         =   "diaSfHrtime.frx":0AB8
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Chkno 
      Height          =   210
      Left            =   7800
      Picture         =   "diaSfHrtime.frx":0E42
      Stretch         =   -1  'True
      Top             =   360
      Visible         =   0   'False
      Width           =   285
   End
End
Attribute VB_Name = "diaSfHrtime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'See the UpdateTables procedure for database revisions
Option Explicit
Dim bOnLoad As Byte
Dim bChanged As Byte
Dim bLoading As Byte

Dim sCustomers(500, 2) As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'Dim RdoQry As rdoQuery
Dim prmObj As ADODB.Parameter
Dim cmdObj As ADODB.Command

Dim bNew  As Boolean

Private Sub cmbEmp_Click()
    Dim strPremNum As String
    Dim strLastName As String
    Dim strFirstName As String
    Dim iPremNum As Long
    
    Dim RdoEmp As ADODB.Recordset
    
    strPremNum = cmbEmp
    
   If (strPremNum = "") Then
      cmbEmp = "ALL"
      strPremNum = "ALL"
   End If
   
   If (strPremNum = "ALL") Then
       lblNme = " - ALL - "
   Else
       
       sSql = "SELECT PREMLSTNAME, PREMFSTNAME,PREMTERMDT FROM EmplTable WHERE PREMNUMBER = " & CLng(strPremNum)
       bSqlRows = clsADOCon.GetDataSet(sSql, RdoEmp, ES_FORWARD)
       If bSqlRows Then
          With RdoEmp
          
           If (Not IsNull(!PREMTERMDT)) Then
              MsgBox "Not a Current Employee.", vbInformation, Caption
              Set RdoEmp = Nothing
              lblNme = "Not a Current Employee"
              Exit Sub
           End If
          
           strLastName = "" & Trim(!PREMLSTNAME)
           strFirstName = "" & Trim(!PREMFSTNAME)
             
           lblNme = strLastName & " " & strFirstName
          End With
          
          cmdUpdate.Enabled = False
          cmdEnd.Enabled = False
       End If
       Set RdoEmp = Nothing
   End If
   
End Sub

Private Sub cmbEmp_LostFocus()
   cmbEmp_Click
End Sub

Private Sub cmbSfCd_LostFocus()
   cmbSfCd_Click
End Sub

Private Sub cmbSfCd_Click()
    Dim RdoSfCd As ADODB.Recordset
    Dim strSfCd As String
    
    If bLoading = 0 Then
        Exit Sub
    End If
        
    strSfCd = cmbSfCd
    If (strSfCd = "") Then
        txtSfBeg = ""
        txtSfEnd = ""
        txtSLBeg = ""
        txtSLEnd = ""
        txtAdjMin = "0"
        
        txtTotHrs = Format(0, "##0.00")
        lblAddRate = Format(0, "##0.00")
    Else
        
        sSql = "SELECT SFSTHR,SFENHR,SFLUNSTHR, SFLUNENHR,SFADDRT,SFADJHR FROM " _
               & " sfcdTable WHERE SFCODE= '" & strSfCd & "'"
        bSqlRows = clsADOCon.GetDataSet(sSql, RdoSfCd, ES_FORWARD)
        If bSqlRows Then
           Dim minutes As Long
           Dim lHrs As Currency
           Dim lLunHrs As Currency
           
           With RdoSfCd
            txtSfBeg = "" & Trim(!SFSTHR)
            txtSfEnd = "" & Trim(!SFENHR)
            txtSLBeg = "" & Trim(!SFLUNSTHR)
            txtSLEnd = "" & Trim(!SFLUNENHR)
            lblAddRate = "" & IIf(IsNull(Trim(!SFADDRT)), "0.00", Trim(!SFADDRT))
            txtAdjMin = "" & IIf(IsNull(Trim(!SFADJHR)), "0.00", Trim(!SFADJHR))
            
            If (Trim(txtSfBeg) <> "" And Trim(txtSfEnd) <> "") Then
               ' if the shift overlaps to next day, append the date and do the additions
               If (CDate(txtSfBeg) > CDate(txtSfEnd)) Then
                  Dim strBegDate As String
                  Dim strEndDate As String

                  strBegDate = Format(Now, "mm/dd/yy ") & txtSfBeg
                  strEndDate = Format(DateAdd("d", 1, Now), "mm/dd/yy ") & txtSfEnd
                  minutes = DateDiff("n", CDate(strBegDate), CDate(strEndDate))
                  lHrs = Format(minutes / 60, "##0.00")
               Else
                  minutes = DateDiff("n", txtSfBeg, txtSfEnd)
                  lHrs = Format(minutes / 60, "##0.00")
               End If
            Else
                lHrs = 0
            End If
            
            If (Trim(txtSLBeg) <> "" And Trim(txtSLEnd) <> "") Then
                minutes = DateDiff("n", txtSLBeg, txtSLEnd)
                lLunHrs = Format(minutes / 60, "##0.00")
            Else
                lLunHrs = 0
            End If
            ' Get the lunch hours
            txtTotHrs = Format((lHrs - lLunHrs), "##0.00")
        
           End With
        End If
    End If
    
End Sub

Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdEnd_Click()
    Dim iList As Integer
    For iList = 1 To Grd.Rows - 1
        Grd.Col = 9
        Grd.row = iList
        ' Only if the part is checked
        If Grd.CellPicture = Chkyes.picture Then
            Set Grd.CellPicture = Chkno.picture
        End If
        
        Grd.Col = 4
        If Grd.CellPicture = Chkyes.picture Then
            Set Grd.CellPicture = Chkno.picture
        End If
        
            Grd.Col = 10
        If Grd.CellPicture = Chkyes.picture Then
            Set Grd.CellPicture = Chkno.picture
        End If
        
    Next
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 2104
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub

Private Sub cmdUpdate_Click()
   Dim iList As Integer
   Dim bApplyOnlyLSC  As Boolean
   Dim bApplySC As Boolean
   Dim bApplyAddlnRate As Boolean
   On Error GoTo DiaErr1
   MouseCursor 13
   Err.Clear
    
   Dim tc As New ClassTimeCharge
   ' Go throught all the record int he grid and re-schedule MO
   For iList = 1 To Grd.Rows - 1
      
      bApplyOnlyLSC = False
      bApplySC = False
      bApplyAddlnRate = False
      
      Grd.Col = 9
      Grd.row = iList
       
      If Grd.CellPicture = Chkyes.picture Then
         bApplySC = True
      End If
      
      Grd.Col = 4
      If Grd.CellPicture = Chkyes.picture Then
         bApplyOnlyLSC = True
      End If
      
      Grd.Col = 10
      If Grd.CellPicture = Chkyes.picture Then
         bApplyAddlnRate = True
      End If
      ' Only if the part is checked
      If ((bApplySC = True) Or (bApplyOnlyLSC = True) Or (bApplyAddlnRate = True)) Then
          Dim minutes As Integer
          Dim strIDname As String
          Dim strSfBeg As String
          Dim strSfEnd As String
          Dim strSfHrs As String
          Dim strTCBeg As String
          Dim strTCEnd As String
          Dim strTCRegHrs As String
          Dim strDate As String
          
          Dim strEmpID, strEmpLName As String
          Dim arrEmp() As String
          
          Grd.Col = 0
          strIDname = Grd.Text
          arrEmp = Split(strIDname, "-")
          strEmpID = arrEmp(0)
          strEmpLName = arrEmp(1)
          ' replace the asterix
          strEmpID = Replace(arrEmp(0), Chr$(42), "")
          strDate = txtDte
          
          Grd.Col = 1
          strTCBeg = Grd.Text
          Grd.Col = 2
          strTCEnd = Grd.Text
          Grd.Col = 3
          strTCRegHrs = Grd.Text
          Grd.Col = 5
          strSfBeg = Grd.Text
          Grd.Col = 6
          strSfEnd = Grd.Text
          Grd.Col = 7
          strSfHrs = Grd.Text
          
          
          ' if we need to apply addtional rates
          If (bApplyAddlnRate = True) Then
             Dim cAddRate As Currency
             cAddRate = CDbl(lblAddRate)
             ApplyAdditionalRate CLng(strEmpID), strDate, cAddRate
          End If
          
          ' if we need to apply addtional rates
          If (bApplyOnlyLSC = True) Then
             
             tc.ApplySCOnlyToLunch CLng(strEmpID), strDate, strTCBeg, strTCEnd, strTCRegHrs, _
                                         strSfBeg, strSfEnd, strSfHrs
          End If
          
          If (bApplySC = True) Then
             tc.ApplyShiftCode CLng(strEmpID), strDate, strTCBeg, strTCEnd, strTCRegHrs, _
                                         strSfBeg, strSfEnd, strSfHrs
         End If
         
      End If
   Next
   
   If Err <> 0 Then
      MsgBox "Couldn't Successfully Update..", _
         vbInformation, Caption
   End If
   
   Dim strSfcode As String
   strDate = txtDte
   strSfcode = cmbSfCd
   FillGrid 0, strDate, strSfcode
   
   MouseCursor 0
   Exit Sub

DiaErr1:
   sProcName = "cmdUpdate"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me

End Sub

Private Sub ApplyAdditionalRate(iEmpNum As Long, strCardDate As String, cAddRt As Currency)
    
    Dim RdoEmp As ADODB.Recordset
    Dim strTCCard As String
    Dim cOHPerc As Currency
    Dim cRate As Currency
    Dim cTotRate As Currency
    On Error GoTo DiaErr1
        
    sSql = "SELECT TMCARD FROM viewShiftCdEmployeeDetail, TchdTable " & _
        "WHERE TMEMP = PREMNUMBER AND " & _
            "PREMNUMBER = '" & CStr(iEmpNum) & "' AND " & _
            "TMDAY = '" & strCardDate & "'"
    
    bSqlRows = clsADOCon.GetDataSet(sSql, RdoEmp, ES_FORWARD)
    
    If bSqlRows Then
        strTCCard = Trim(RdoEmp!TMCARD)
        ClearResultSet RdoEmp
    
         ' get current overhead rate
         GetOverHeadPerct iEmpNum, strTCCard, cRate, cOHPerc
         
         
         sSql = "UPDATE TcitTable SET TCRATE = (TCRATE + " & CStr(cAddRt) & ") " & _
             " WHERE TCCARD = '" & strTCCard & "' AND TCEMP = '" & CStr(iEmpNum) & "'"
         
         clsADOCon.ExecuteSql sSql ', rdExecDirect
         
         cTotRate = ((CInt(cRate) + CInt(cAddRt)) * CInt(cOHPerc)) / 100
         sSql = "UPDATE TcitTable SET TCOHRATE = " & CStr(cTotRate) & _
             " WHERE TCCARD = '" & strTCCard & "' AND TCEMP = '" & CStr(iEmpNum) & "' AND TCOHRATE <> 0"
         
         clsADOCon.ExecuteSql sSql ', rdExecDirect
         
         
'         sSql = "UPDATE TcitTable SET TCRATE = (TCRATE + " & CStr(cAddRt) & ") " & _
'             " WHERE TCCARD = '" & strTCCard & "' AND TCEMP = '" & CStr(iEmpNum) & "'"
'         clsADOCon.ExecuteSQL sSql ', rdExecDirect
         
    End If
    
    Set RdoEmp = Nothing
    
    Exit Sub
    
DiaErr1:
   sProcName = "ApplyAdditionalRate"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
       
End Sub


Private Sub GetOverHeadPerct(ByVal iEmpNum As Long, ByVal strTCCard As String, ByRef cRate As Currency, ByRef cOHPerc As Currency)
    
   Dim RdoOH As ADODB.Recordset
   Dim cOHrate As Currency
   
   On Error GoTo DiaErr1
        
   sSql = "SELECT DISTINCT ISNULL(TCRATE, 0) TCRATE, ISNULL(TCOHRATE, 0) TCOHRATE FROM TcitTable " & _
       "WHERE TCEMP = '" & CStr(iEmpNum) & "' AND " & _
           "TCCARD = '" & strTCCard & "' AND TCRATE <> 0 AND TCOHRATE <> 0"
    
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoOH, ES_FORWARD)
    
   If bSqlRows Then
      cRate = Trim(RdoOH!TCRATE)
      cOHrate = Trim(RdoOH!TCOHRATE)
      
      If (cRate <> 0) Then
         cOHPerc = (cOHrate * 100) / cRate
      Else
         cOHPerc = 0
      End If
      ClearResultSet RdoOH
   Else
      cOHPerc = 0
   End If
   
   Set RdoOH = Nothing
   
   Exit Sub
   
DiaErr1:
   sProcName = "GetOverHeadPerct"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
       
End Sub
Private Sub Form_Activate()
    Dim iRow As String
    Dim strDate As String
    Dim strSfcode As String
    bLoading = 0
    
    If bOnLoad = 1 Then
        cmbEmp.AddItem ("ALL")
        sSql = "select PREMNUMBER, PREMLSTNAME, PREMFSTNAME from EmplTable where PREMTERMDT IS NULL AND PREMSTATUS <> 'D' " _
               & "order by PREMNUMBER"
        LoadNumComboBox cmbEmp, "00000"
        If bSqlRows Then
            cmbEmp = cmbEmp.List(0)
            lblNme = " - ALL - "
        End If
        strDate = txtDte
        
        sSql = "select SFCODE FROM sfcdTable ORDER BY SFCODE"
        LoadNumComboBox cmbSfCd, "00"
        If bSqlRows Then
            cmbSfCd = ""
        End If
        
        strSfcode = cmbSfCd
        iRow = FillGrid(0, strDate, strSfcode)
        bOnLoad = 1
    End If
    ' Loaded completed
    bLoading = 1
    MouseCursor 0
   
End Sub

Private Sub Form_Load()
   
   
   FormLoad Me
   FormatControls
   
   sSql = "SELECT * FROM EmplTable WHERE PREMNUMBER = ? "
   
   'Set RdoQry = RdoCon.CreateQuery("", sSql)
   'RdoQry.MaxRows = 1
   Set cmdObj = New ADODB.Command
   cmdObj.CommandText = sSql

   Set prmObj = New ADODB.Parameter
   prmObj.Type = adInteger
   cmdObj.Parameters.Append prmObj
   
   txtDte = Format(Now - 1, "mm/dd/yy")
   If sCurrDate = "" Then
      If Format(txtDte, "w") = 1 Then
         txtDte = Format(Now - 2, "mm/dd/yy")
      End If
   Else
      txtDte = sCurrDate
   End If
   GetWeekEnd
   'GetOptions
   bOnLoad = 1
   
   With Grd
      .ColAlignment(0) = 0
      .ColAlignment(1) = 1
      .ColAlignment(2) = 1
      .ColAlignment(3) = 1
      .ColAlignment(4) = 1
      .ColAlignment(5) = 1
      .ColAlignment(6) = 1
      .ColAlignment(7) = 1
      .ColAlignment(8) = 2
      .ColAlignment(9) = 2
      .Rows = 1
      .row = 0
      .Col = 0
      .Text = "Employee (Number)"
      .Col = 1
      .Text = "TC Start"
      .Col = 2
      .Text = "TC End"
      .Col = 3
      .Text = "TOT Hrs"
      .Col = 4
      .Text = "Only Lunch SC"
      .Col = 5
      .Text = "SC Start"
      .Col = 6
      .Text = "SC End"
      .Col = 7
      .Text = "TOT Hrs"
      .Col = 8
      .Text = "LUNCH"
      .Col = 9
      .Text = "Apply SC"
      .Col = 10
      .Text = "Addnl Rate"
      
      .ColWidth(0) = 2200
      .ColWidth(1) = 900
      .ColWidth(2) = 900
      .ColWidth(3) = 900
      .ColWidth(4) = 1200
      .ColWidth(5) = 900
      .ColWidth(6) = 900
      .ColWidth(7) = 900
      .ColWidth(8) = 900
      .ColWidth(9) = 1000
      .ColWidth(10) = 1000
      
   End With
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set prmObj = Nothing
   Set cmdObj = Nothing
   Set diaSfcode = Nothing
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Function FillGrid(lEmpNum As Long, strDate As String, strSfcode As String) As Integer
    Dim RdoGrd As ADODB.Recordset
    Dim strEmp As String
    
    On Error Resume Next
    Grd.Rows = 1
    On Error GoTo DiaErr1
    
    If (lEmpNum = 0) Then
        strEmp = "%"
    Else
        strEmp = CStr(lEmpNum)
    End If
    
'    sSql = "SELECT SFCODE, PREMNUMBER, PREMLSTNAME, SFSTHR, SFENHR, ISNULL(SFADJHR, 0) SFADJHR," & _
'            "ISNULL(SFLUNSTHR, 0) SFLUNSTHR, ISNULL(SFLUNENHR, 0) SFLUNENHR, TMSTART, TMSTOP, TMREGHRS " & _
'         "From viewShiftCdEmployeeDetail, TchdTable " & _
'        "WHERE TMEMP = PREMNUMBER AND " & _
'            "TMDAY = '" & strDate & "' AND TMEMP LIKE '" & strEmp & "' AND " & _
'            " TMSTART <> '' AND TMSTOP <> '' " & _
'            " AND SFCODE = '" & strSfCode & "' ORDER BY PREMNUMBER"

    sSql = "SELECT c.SFCODE, PREMNUMBER, PREMLSTNAME, c.SFSTHR, c.SFENHR, ISNULL(c.SFADJHR, 0) SFADJHR," & _
            "ISNULL(c.SFLUNSTHR, 0) SFLUNSTHR, ISNULL(c.SFLUNENHR, 0) SFLUNENHR, TMSTART, TMSTOP, TMREGHRS, " & _
         "ISNULL(TMSFCODEAPPLD, 0) TMSFCODEAPPLD " & _
         " FROM viewShiftCdEmployeeDetail b, TchdTable, SfcdTable c " & _
                 "WHERE c.SFREF = '" & strSfcode & "'  AND TMEMP = PREMNUMBER AND " & _
            "TMDAY = '" & strDate & "' AND TMEMP LIKE '" & strEmp & "' AND " & _
            " TMSTART <> '' AND TMSTOP <> '' " & _
            " AND " & vbCrLf _
            & "(SELECT DISTINCT " & vbCrLf _
                     & "(CASE DATEPART(weekday, '" & strDate & "')" & vbCrLf _
                     & "WHEN 1 THEN ISNULL(SFREFSUN, SFREF)" & vbCrLf _
                     & "WHEN  2 THEN ISNULL(SFREFMON, SFREF)" & vbCrLf _
                     & "WHEN  3 THEN ISNULL(SFREFTUE, SFREF)" & vbCrLf _
                     & "WHEN  4 THEN ISNULL(SFREFWED, SFREF)" & vbCrLf _
                     & "WHEN  5 THEN ISNULL(SFREFTHU, SFREF)" & vbCrLf _
                     & "WHEN  6 THEN ISNULL(SFREFFRI, SFREF)" & vbCrLf _
                     & "WHEN  7 THEN ISNULL(SFREFSAT, SFREF)" & vbCrLf _
                     & "Else SFREF" & vbCrLf _
               & "END)" & vbCrLf _
               & "FROM sfempTable a WHERE a.PREMNUMBER = b.PREMNUMBER)" & vbCrLf _
               & " = '" & strSfcode & "' ORDER BY PREMNUMBER"
    
    Debug.Print sSql
    
    Dim tc As New ClassTimeCharge
    Dim lSftMin As Long
    Dim strBeg As String
    Dim strEnd As String
    Dim strSfBeg As String
    Dim strSfEnd As String
    Dim strSfHrs As String
    Dim strLBeg As String
    Dim strLEnd As String
    Dim strLunHrs As String
    Dim strAdjHrs As String
    
    
    bSqlRows = clsADOCon.GetDataSet(sSql, RdoGrd, ES_FORWARD)
    If bSqlRows Then
        With RdoGrd
            Do Until .EOF
            Grd.Rows = Grd.Rows + 1
            Grd.row = Grd.Rows - 1
            Grd.Col = 0
            Grd.Text = IIf(CInt(!TMSFCODEAPPLD) = 1, "*", "") & Trim(!PREMNUMBER) & " - " & Trim(!PREMLSTNAME)
            Grd.Col = 1
            strBeg = Trim(!TMSTART)
            Grd.Text = "" & strBeg
            Grd.Col = 2
            strEnd = Trim(!TMSTOP)
            Grd.Text = "" & strEnd
            Grd.Col = 3
            Grd.Text = "" & Trim(!TMREGHRS)
            
            Grd.Col = 4
            Set Grd.CellPicture = Chkno.picture
            
            strSfBeg = "" & Trim(!SFSTHR)
            strSfEnd = "" & Trim(!SFENHR)
            strLBeg = "" & Trim(!SFLUNSTHR)
            strLEnd = "" & Trim(!SFLUNENHR)
            strAdjHrs = "" & Trim(!SFADJHR)
            
            Dim strBegDate As String
            Dim strEndDate As String
            Dim strLunBegDate As String
            Dim strLunEndDate As String
            Dim strSfBegDate As String
            Dim strSfEndDate As String
            
            
            If (CDate(strSfBeg) > CDate(strSfEnd)) Then
               strBegDate = Format(Now, "mm/dd/yy ") & strBeg
               strEndDate = Format(DateAdd("d", 1, Now), "mm/dd/yy ") & strEnd
               strSfBegDate = Format(Now, "mm/dd/yy ") & strSfBeg
               strSfEndDate = Format(DateAdd("d", 1, Now), "mm/dd/yy ") & strSfEnd
            Else
               strBegDate = Format(Now, "mm/dd/yy ") & strBeg
               If (CDate(strBeg) > CDate(strEnd)) Then
                  strEndDate = Format(DateAdd("d", 1, Now), "mm/dd/yy ") & strEnd
               Else
                  strEndDate = Format(Now, "mm/dd/yy ") & strEnd
               End If
               
               strSfBegDate = Format(Now, "mm/dd/yy ") & strSfBeg
               strSfEndDate = Format(Now, "mm/dd/yy ") & strSfEnd
            End If
               
            'strLunBegDate = Format(Now, "mm/dd/yy ") & strLBeg
            'strLunEndDate = Format(Now, "mm/dd/yy ") & strLEnd
            
            If (strLBeg <> "") Then
               If (CDate(strSfBeg) > CDate(strLBeg)) Then
                  strLunBegDate = Format(DateAdd("d", 1, Now), "mm/dd/yy ") & strLBeg
                  strLunEndDate = Format(DateAdd("d", 1, Now), "mm/dd/yy ") & strLEnd
               Else
                  strLunBegDate = Format(Now, "mm/dd/yy ") & strLBeg
                  strLunEndDate = Format(Now, "mm/dd/yy ") & strLEnd
               End If
            End If
            
            ' Adjust the checked shift times
            tc.AdjustShiftStartEndTime strBegDate, strEndDate, strSfBegDate, strSfEndDate, strAdjHrs
        
            ' Adjust the checked shift times and lunch time
            Dim lLunMin As Long
            Dim strTmpBeg As String
            Dim strTmpEnd As String
            'tc.AdjWithLunchStartEndTime strBeg, strEnd, lLunMin, strLBeg, strLEnd
            'tc.AdjWithLunchStartEndTime strBegDate, strEndDate, lLunMin, strLunBegDate, strLunEndDate
            If (strLBeg <> "") Then
               tc.AdjWithLunchStartEndTime strBegDate, strEndDate, lLunMin, strLunBegDate, strLunEndDate
            End If
            
            strTmpBeg = tc.GetTime(strBegDate)
            Grd.Col = 5
            'Grd.Text = "" & tc.RoundMinutes(strTmpBeg)
            Grd.Text = "" & tc.GetTime(strBegDate)
            
            strTmpEnd = tc.GetTime(strEndDate)
            Grd.Col = 6
            'Grd.Text = "" & tc.RoundMinutes(strTmpEnd)
            Grd.Text = "" & tc.GetTime(strEndDate)
            
            If (tc.IsValidTime(strBegDate) And _
                        tc.IsValidTime(strEndDate)) Then
                lSftMin = DateDiff("n", strBegDate, strEndDate)
                strSfHrs = Format((lSftMin - lLunMin) / 60, "##0.00")
            Else
                strSfHrs = Format(0, "##0.00")
            End If
            
            Grd.Col = 7
            Grd.Text = strSfHrs
            
            ' Calculate the lunch hours
            strLBeg = "" & Trim(!SFLUNSTHR)
            strLEnd = Trim(!SFLUNENHR)
            
            strLunHrs = Format(lLunMin / 60, "##0.00")
            
            Grd.Col = 8
            Grd.Text = "" & strLunHrs
            
            Grd.Col = 9
            Set Grd.CellPicture = Chkno.picture

            Grd.Col = 10
            Set Grd.CellPicture = Chkno.picture
            
            .MoveNext
         Loop
         ClearResultSet RdoGrd
      End With
   End If
   Set RdoGrd = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "fillgrid"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub Grd_Click()
   Dim row As Integer
   row = Grd.mouseRow      ' row changes when you referene grid
   
   If Grd.Col = 4 Or Grd.Col = 9 Or Grd.Col = 10 Then
      ' if heading selected, set all rows to opposite of first row
      If (Grd.Rows > 1) Then
         If row = 0 Then
            Dim pic As Image
            Grd.row = 1
            If Grd.CellPicture = Chkyes.picture Then
               Set pic = Chkno
            Else
               Set pic = Chkyes
            End If
            
            For row = 1 To Grd.Rows - 1
               Grd.row = row
               Set Grd.CellPicture = pic.picture
            Next row
            
         Else
            If Grd.CellPicture = Chkyes.picture Then
               Set Grd.CellPicture = Chkno.picture
            Else
               Set Grd.CellPicture = Chkyes.picture
            End If
         End If

      End If
   End If


   
End Sub


Private Sub GetWeekEnd()
   Dim RdoGet As ADODB.Recordset
   Dim A As Integer
   Dim iList As Integer
   Dim dDate As Date
   Dim sWeekEnds As String
   
   On Error GoTo DiaErr1
   sSql = "SELECT WEEKENDS FROM ComnTable WHERE COREF=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoGet, ES_FORWARD)
   If bSqlRows Then
      With RdoGet
         sWeekEnds = "" & Trim(!WEEKENDS)
         .Cancel
      End With
      If sWeekEnds = "Sat" Then iList = 7 Else iList = 8
   End If
   dDate = txtDte
   A = Format(txtDte, "w")
   lblWen = Format(dDate + (iList - A), "mm/dd/yy")
   Set RdoGet = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getweeken"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub txtDte_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtDte_LostFocus()
    Dim iRow As String
    Dim strDate As String
    Dim strEmp As String
    Dim strSfcode As String
    
    strDate = txtDte
    strEmp = cmbEmp
    If (strEmp = "ALL") Then
        strEmp = 0
    End If
    
    strSfcode = cmbSfCd
    iRow = FillGrid(CLng(strEmp), strDate, strSfcode)

    cmdUpdate.Enabled = True
    cmdEnd.Enabled = True
End Sub

