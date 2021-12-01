VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form diaSfHrAdmtime 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Global Time Charge For Employees"
   ClientHeight    =   8460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12015
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "diaSfHrAdmtime.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8460
   ScaleWidth      =   12015
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbEAct 
      Height          =   315
      Left            =   6360
      TabIndex        =   3
      Tag             =   "2"
      Top             =   240
      Width           =   1455
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
      Height          =   480
      Left            =   2400
      TabIndex        =   38
      ToolTipText     =   " Clear the selection"
      Top             =   2400
      Width           =   1920
   End
   Begin VB.CommandButton CmdSelAll 
      Caption         =   "Selection All"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   120
      TabIndex        =   37
      ToolTipText     =   " Select All"
      Top             =   2400
      Width           =   1920
   End
   Begin VB.ComboBox cmbTCode 
      Height          =   315
      Left            =   5280
      TabIndex        =   35
      Tag             =   "3"
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton cmdSel 
      Caption         =   "&Select"
      Enabled         =   0   'False
      Height          =   315
      Left            =   6480
      TabIndex        =   11
      TabStop         =   0   'False
      ToolTipText     =   "Enter Updated Time Card"
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox txtSLBeg 
      Height          =   315
      Left            =   1320
      TabIndex        =   8
      Text            =   "  :"
      Top             =   1800
      Width           =   735
   End
   Begin VB.TextBox txtAccNum 
      Height          =   285
      Left            =   6000
      TabIndex        =   5
      Top             =   720
      Width           =   1815
   End
   Begin MSFlexGridLib.MSFlexGrid Grd 
      Height          =   4575
      Left            =   120
      TabIndex        =   15
      ToolTipText     =   "Click To Select Or Scroll And Press Enter (Also Page Up And Page Down)"
      Top             =   3000
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   8070
      _Version        =   393216
      Rows            =   3
      Cols            =   6
      FixedRows       =   2
      FixedCols       =   0
      FocusRect       =   2
      ScrollBars      =   2
   End
   Begin VB.TextBox txtAdjMin 
      Height          =   315
      Left            =   3960
      TabIndex        =   31
      Top             =   7920
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CheckBox chkAddRate 
      Height          =   255
      Left            =   1200
      TabIndex        =   27
      Top             =   8760
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtSfBeg 
      Height          =   315
      Left            =   1320
      TabIndex        =   6
      Text            =   "  :"
      Top             =   1320
      Width           =   735
   End
   Begin VB.TextBox txtSfEnd 
      Height          =   315
      Left            =   3360
      TabIndex        =   7
      Text            =   "  :"
      Top             =   1320
      Width           =   735
   End
   Begin VB.TextBox txtTotHrs 
      Height          =   315
      Left            =   7560
      TabIndex        =   10
      Top             =   7560
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtSLEnd 
      Height          =   315
      Left            =   3360
      TabIndex        =   9
      Top             =   1800
      Width           =   735
   End
   Begin VB.ComboBox cmbSfCd 
      Height          =   315
      Left            =   3480
      TabIndex        =   2
      Tag             =   "2"
      Top             =   190
      Width           =   1095
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      Enabled         =   0   'False
      Height          =   315
      Left            =   8040
      TabIndex        =   12
      TabStop         =   0   'False
      ToolTipText     =   "Enter Updated Time Card"
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   "C&lear"
      Enabled         =   0   'False
      Height          =   315
      Left            =   9600
      TabIndex        =   13
      TabStop         =   0   'False
      ToolTipText     =   "Cancel Time Card Entry"
      Top             =   1800
      Visible         =   0   'False
      Width           =   875
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   10680
      TabIndex        =   14
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
      Left            =   9120
      TabIndex        =   4
      Tag             =   "3"
      Top             =   270
      Width           =   1095
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   11400
      Picture         =   "diaSfHrAdmtime.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   120
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   250
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Emp Account"
      Height          =   255
      Index           =   15
      Left            =   5280
      TabIndex        =   39
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Type Code"
      Height          =   255
      Index           =   14
      Left            =   4440
      TabIndex        =   36
      Top             =   1365
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Account Number"
      Height          =   255
      Index           =   13
      Left            =   4560
      TabIndex        =   34
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "* Applied Employee Shift code"
      Height          =   255
      Index           =   11
      Left            =   480
      TabIndex        =   33
      Top             =   7920
      Width           =   3135
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Time Adjustment in minutes"
      Height          =   375
      Index           =   10
      Left            =   6480
      TabIndex        =   32
      Top             =   8160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   """TOT Hrs"" doesn't include Lunch"
      Height          =   255
      Index           =   9
      Left            =   4920
      TabIndex        =   30
      ToolTipText     =   "Additional rate"
      Top             =   8280
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label lblAddRate 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3240
      TabIndex        =   29
      Top             =   8760
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Additional Rate"
      Height          =   255
      Index           =   8
      Left            =   1560
      TabIndex        =   28
      ToolTipText     =   "Additional rate"
      Top             =   8760
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Start Time"
      Height          =   255
      Index           =   5
      Left            =   480
      TabIndex        =   26
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "End Time"
      Height          =   255
      Index           =   4
      Left            =   2520
      TabIndex        =   25
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "LunchEnd Time"
      Height          =   255
      Index           =   6
      Left            =   2160
      TabIndex        =   24
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "LunchStart Time"
      Height          =   255
      Index           =   7
      Left            =   0
      TabIndex        =   23
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Shift Code"
      Height          =   255
      Index           =   3
      Left            =   2640
      TabIndex        =   22
      Top             =   240
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Employee"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   21
      Top             =   195
      Width           =   735
   End
   Begin VB.Label lblNme 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   960
      TabIndex        =   20
      Top             =   645
      Width           =   2415
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   19
      Top             =   645
      Width           =   615
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Card Date"
      Height          =   255
      Index           =   1
      Left            =   8280
      TabIndex        =   18
      Top             =   315
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Week Ending"
      Height          =   255
      Index           =   12
      Left            =   7200
      TabIndex        =   17
      Top             =   7845
      Width           =   1215
   End
   Begin VB.Label lblWen 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   8400
      TabIndex        =   16
      ToolTipText     =   "Week Ending (System Administration Setup)"
      Top             =   7845
      Width           =   855
   End
   Begin VB.Image Chkyes 
      Height          =   210
      Left            =   10680
      Picture         =   "diaSfHrAdmtime.frx":0AB8
      Stretch         =   -1  'True
      Top             =   1080
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Chkno 
      Height          =   210
      Left            =   10680
      Picture         =   "diaSfHrAdmtime.frx":0E42
      Stretch         =   -1  'True
      Top             =   1320
      Visible         =   0   'False
      Width           =   285
   End
End
Attribute VB_Name = "diaSfHrAdmtime"
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
       ' txtSLBeg = ""
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
        Grd.Col = 5
        If Grd.CellPicture = Chkyes.Picture Then
            Set Grd.CellPicture = Chkno.Picture
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

Private Sub cmdSel_Click()
    
   Dim iRow As String
   Dim strDate As String
   Dim strEmp As String
   Dim strSfcode As String
   
   strDate = txtDte
   strEmp = cmbEmp
   If (strEmp = "ALL") Then
       strEmp = 0
   End If
   
   cmdUpdate.Enabled = True
   cmdEnd.Enabled = True
   
   strSfcode = cmbSfCd
   
   If (Trim(txtSfBeg.Text) = "" Or Trim(txtSfEnd.Text) = "") Then
      MsgBox "Please Enter the Start and End time.", vbExclamation, Caption
      Exit Sub
   End If
   iRow = FillGrid(CLng(strEmp), strSfcode)

End Sub

Private Sub cmdSelAll_Click()
   
   Dim iList As Integer
   For iList = 1 To Grd.Rows - 1
       Grd.Col = 5
       Grd.Row = iList
       ' Only if the part is checked
       If Grd.CellPicture = Chkno.Picture Then
           Set Grd.CellPicture = Chkyes.Picture
       End If
   Next
End Sub

Private Sub cmdClear_Click()
    Dim iList As Integer
    For iList = 1 To Grd.Rows - 1
        Grd.Col = 5
        Grd.Row = iList
        ' Only if the part is checked
        If Grd.CellPicture = Chkyes.Picture Then
            Set Grd.CellPicture = Chkno.Picture
        End If
    Next
End Sub


Private Sub cmdUpdate_Click()
   Dim iList As Integer
   Dim bAddTC As Boolean
   Dim strAccount As String
   
   On Error GoTo DiaErr1
   MouseCursor 13
   Err.Clear
    
   
   clsADOCon.ADOErrNum = 0
   clsADOCon.BeginTrans

   
   Dim rdo As ADODB.Recordset
   Dim strTimeCode As String
'   sSql = "select TYPECODE from TmcdTable where TYPETYPE = 'R' ORDER BY TYPESEQ"
'   If GetDataSet(rdo) Then
'      regularTimeCode = rdo.rdoColumns(0)
'   Else
'      regularTimeCode = "RT"
'   End If
'   Set rdo = Nothing
   
   strTimeCode = cmbTCode.Text
   If (Trim(strTimeCode) = "") Then strTimeCode = "RT"
   
   strAccount = txtAccNum.Text
   
   If (Trim(strAccount) = "") Then
      MsgBox "Pelase enter the Indirect Account number", vbExclamation, Caption
      Exit Sub
   End If
   
   Dim tc As New ClassTimeCharge
   ' Go throught all the record int he grid and re-schedule MO
   For iList = 1 To Grd.Rows - 1
      
      Grd.Row = iList
      Grd.Col = 5
      If Grd.CellPicture = Chkyes.Picture Then
         Dim minutes As Integer
         Dim strIDname As String
         Dim strSfBeg As String
         Dim strSfEnd As String
         Dim strSfHrs As String
         Dim strSLBeg As String
         Dim strSLEnd As String
         Dim strDate As String
         Dim StartDateTime As Variant
         Dim EndDateTime As Variant
         Dim SLStartDateTime As Variant
         Dim SLEndDateTime As Variant
         Dim strLunHrs As String
         Dim strEmpID, strEmpLName As String
         Dim arrEmp() As String
         Dim sNewCard  As String
         
         Grd.Col = 0
         strIDname = Grd.Text
         arrEmp = Split(strIDname, "-")
         ' replace the asterix
         strEmpID = Replace(arrEmp(0), Chr$(42), "")
         strEmpLName = arrEmp(1)
         strDate = txtDte
         
         Grd.Col = 1
         strSfBeg = Grd.Text
         Grd.Col = 2
         strSfEnd = Grd.Text
         Grd.Col = 3
         strLunHrs = Grd.Text
         Grd.Col = 4
         strSfHrs = Grd.Text
         
         strSLBeg = txtSLBeg
         strSLEnd = txtSLEnd
         
         Sleep (300)
         sNewCard = GetNewNumber(CLng(strEmpID))

         sSql = "INSERT INTO TchdTable (TMCARD,TMEMP,TMDATE,TMDAY," _
             & "TMWEEK, TMSTART, TMSTOP, TMREGHRS) " _
             & "VALUES('" & sNewCard & "'," & strEmpID & ",'" _
             & Format(txtDte, "mm/dd/yy") & "','" & txtDte & "'" _
             & ",'" & lblWen & "','" & strSfBeg & "','" _
             & strSfEnd & "','" & strSfHrs & "')"
         clsADOCon.ExecuteSql sSql 'rdExecDirect
         
         
         If (CDate(strSfBeg) > CDate(strSfEnd)) Then
            StartDateTime = Format(strDate, "mm/dd/yy ") & strSfBeg & "m"
            EndDateTime = Format(DateAdd("d", 1, strDate), "mm/dd/yy ") & strSfEnd & "m"
         Else
            StartDateTime = Format(strDate, "mm/dd/yy ") & strSfBeg & "m"
            EndDateTime = Format(strDate, "mm/dd/yy ") & strSfEnd & "m"
         End If
         
'         StartDateTime = strDate & " " & strSfBeg & "m"
'         EndDateTime = strDate & " " & strSfEnd & "m"

         If (CDate(strSfBeg) > CDate(strSLBeg)) Then
            SLStartDateTime = Format(DateAdd("d", 1, strDate), "mm/dd/yy ") & strSLBeg & "m"
            SLEndDateTime = Format(DateAdd("d", 1, strDate), "mm/dd/yy ") & strSLEnd & "m"
         Else
            SLStartDateTime = Format(strDate, "mm/dd/yy ") & strSLBeg & "m"
            SLEndDateTime = Format(strDate, "mm/dd/yy ") & strSLEnd & "m"
         End If
                  
         'StartDateTime = strDate & " " & strSfBeg & "m"
         'EndDateTime = strDate & " " & strSfEnd & "m"

         'SLStartDateTime = strDate & " " & strSLBeg & "m"
         'SLEndDateTime = strDate & " " & strSLEnd & "m"
         
         If (strSLBeg <> "") And (strSLEnd <> "") Then
         
            ' shift start to Lunch start
            tc.CreateTimeCharge sNewCard, CLng(strEmpID), StartDateTime, SLStartDateTime, _
               strTimeCode, "I", strAccount, "", 0, 0, 0, "TS", 0, 0, 0, ""
            
            ' shift Lunch end to End
            tc.CreateTimeCharge sNewCard, CLng(strEmpID), SLEndDateTime, EndDateTime, _
               strTimeCode, "I", strAccount, "", 0, 0, 0, "TS", 0, 0, 0, ""
         
         Else
            tc.CreateTimeCharge sNewCard, CLng(strEmpID), StartDateTime, EndDateTime, _
               strTimeCode, "I", strAccount, "", 0, 0, 0, "TS", 0, 0, 0, ""
         End If

      End If
   Next
   
   If clsADOCon.ADOErrNum = 0 Then
      clsADOCon.CommitTrans
      MsgBox "Successfully Update Time..", _
         vbInformation, Caption
   Else
      clsADOCon.RollbackTrans
      MsgBox "Couldn't add Time Cards.", vbExclamation, Caption
   
   End If
   
   Dim strSfcode As String
   strDate = txtDte
   strSfcode = cmbSfCd
   FillGrid 0, strSfcode

   MouseCursor 0
   Exit Sub

DiaErr1:
   sProcName = "cmdUpdate"
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
        
        sSql = "select TYPECODE from TmcdTable ORDER BY TYPESEQ"
        LoadNumComboBox cmbTCode, "00"
        If bSqlRows Then
            cmbTCode = ""
        End If
        
        sSql = "select DISTINCT PREMACCTS from EmplTable ORDER BY PREMACCTS"
        LoadNumComboBox cmbEAct, ""
        If bSqlRows Then
            cmbEAct = ""
        End If
        
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
   
   Set cmdObj = New ADODB.Command
   cmdObj.CommandText = sSql

   Set prmObj = New ADODB.Parameter
   prmObj.Type = adInteger
   cmdObj.Parameters.Append prmObj
   
   'Set RdoQry = RdoCon.CreateQuery("", sSql)
   'RdoQry.MaxRows = 1
   
   txtDte = Format(Now - 1, "mm/dd/yy")
   If sCurrDate = "" Then
      If Format(txtDte, "w") = 1 Then
         txtDte = Format(Now - 2, "mm/dd/yy")
      End If
   Else
      txtDte = sCurrDate
   End If
   'GetOptions
   bOnLoad = 1
   
   With Grd
      .ColAlignment(0) = 0
      .ColAlignment(1) = 1
      .ColAlignment(2) = 1
      .ColAlignment(3) = 1
      .ColAlignment(4) = 1
      .ColAlignment(5) = 1
      .Rows = 1
      .Row = 0
      .Col = 0
      .Text = "Employee (Number)"
      .Col = 1
      .Text = "TC Start"
      .Col = 2
      .Text = "TC End"
      .Col = 3
      .Text = "Lunch Hrs"
      .Col = 4
      .Text = "TOT Hrs"
      .Col = 5
      .Text = "Apply Time"
      
      .ColWidth(0) = 2200
      .ColWidth(1) = 900
      .ColWidth(2) = 900
      .ColWidth(3) = 900
      .ColWidth(4) = 900
      .ColWidth(5) = 1200
      
   End With
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set diaSfcode = Nothing
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Function FillGrid(lEmpNum As Long, strSfcode As String) As Integer
    Dim RdoGrd As ADODB.Recordset
    Dim strEmp As String
    Dim strEmpAcct As String
    
    On Error Resume Next
    Grd.Rows = 1
    On Error GoTo DiaErr1
    
    strEmpAcct = Trim(cmbEAct)
    
    If (lEmpNum = 0) Then
        strEmp = "%"
    Else
        strEmp = CStr(lEmpNum)
    End If

    Dim lSftMin As Long
    Dim lLunMin As Long
    Dim lTotMin As Long
    Dim strBeg As String
    Dim strEnd As String
    Dim strSfBeg As String
    Dim strSfEnd As String
    Dim strSfHrs As String
    Dim strLBeg As String
    Dim strLEnd As String
    Dim strLunHrs As String
    Dim strAdjHrs As String

    If (Trim(strSfcode) <> "") Then
    
      sSql = "SELECT SFCODE, a.PREMNUMBER, a.PREMLSTNAME, SFSTHR, SFENHR, SFLUNSTHR, SFLUNENHR,SFADJHR" & _
           "   From viewShiftCdEmployeeDetail a, EmplTable b" & _
           "   WHERE a.PREMNUMBER LIKE '" & strEmp & "'" & _
           " AND a.PREMNUMBER = b.PREMNUMBER" & _
           " AND (  (PREMTERMDT IS NULL) or (PREMTERMDT IS NOT NULL AND PREMREHIREDT > PREMTERMDT) )" & _
           " AND ( PREMSTATUS NOT IN ('D','I'))" & _
           " AND SFCODE = '" & strSfcode & "' " & _
           " AND PREMACCTS LIKE '" & strEmpAcct & "%' ORDER BY a.PREMNUMBER"

'      sSql = "SELECT SFCODE, PREMNUMBER, PREMLSTNAME, SFSTHR, SFENHR, SFLUNSTHR, SFLUNENHR,SFADJHR " & _
'           "From viewShiftCdEmployeeDetail " & _
'          "WHERE PREMNUMBER LIKE '" & strEmp & "' AND SFCODE = '" & strSfcode & "' ORDER BY PREMNUMBER"
   Else
      
      sSql = "SELECT SFCODE, a.PREMNUMBER, a.PREMLSTNAME, SFSTHR, SFENHR, SFLUNSTHR, SFLUNENHR,SFADJHR" & _
           "   From viewShiftCdEmployeeDetail a, EmplTable b" & _
           "   WHERE a.PREMNUMBER LIKE '" & strEmp & "'" & _
           " AND a.PREMNUMBER = b.PREMNUMBER" & _
           " AND (  (PREMTERMDT IS NULL) or (PREMTERMDT IS NOT NULL AND PREMREHIREDT > PREMTERMDT) )" & _
           " AND ( PREMSTATUS NOT IN ('D','I')) " & _
           " AND PREMACCTS LIKE '" & strEmpAcct & "%' ORDER BY a.PREMNUMBER"
      
'      sSql = "SELECT SFCODE, PREMNUMBER, PREMLSTNAME, SFSTHR, SFENHR, SFLUNSTHR, SFLUNENHR,SFADJHR " & _
'           "From viewShiftCdEmployeeDetail " & _
'          "WHERE PREMNUMBER LIKE '" & strEmp & "' ORDER BY PREMNUMBER"
   End If
   
   Debug.Print sSql
   
    Dim bTCEntry As Boolean
    Dim strmark As String
   
    bSqlRows = clsADOCon.GetDataSet(sSql, RdoGrd, ES_FORWARD)
    If bSqlRows Then
        With RdoGrd
            Do Until .EOF
            
            ' check for data entry
            bTCEntry = GetTCEntryRecord(Trim(!PREMNUMBER), txtDte.Text)
            
            strmark = ""
            If (bTCEntry = True) Then strmark = "*"
            Grd.Rows = Grd.Rows + 1
            Grd.Row = Grd.Rows - 1
            Grd.Col = 0
            Grd.Text = strmark & Trim(!PREMNUMBER) & " - " & Trim(!PREMLSTNAME)
            Grd.Col = 1
            strSfBeg = txtSfBeg
            Grd.Text = "" & strSfBeg
            Grd.Col = 2
            strSfEnd = txtSfEnd
            Grd.Text = "" & strSfEnd
            
            ' Calculate the lunch hours
            strLBeg = IIf(Trim(txtSLBeg) = "", "0", txtSLBeg)
            strLEnd = IIf(Trim(txtSLEnd) = "", "0", txtSLEnd)
            
            Grd.Col = 3
            If (strLBeg <> "0" And strLEnd <> "0") Then
               lLunMin = DateDiff("n", strLBeg, strLEnd)
               strLunHrs = Format(lLunMin / 60, "##0.00")
            Else
               lLunMin = 0
               strLunHrs = "0"
            End If
            
            Grd.Text = "" & strLunHrs
            
            'lSftMin = DateDiff("n", strSfBeg, strSfEnd)
            Dim strSfBegDate As String
            Dim strSfEndDate As String
            
            
            If (CDate(strSfBeg) > CDate(strSfEnd)) Then
               strSfBegDate = Format(Now, "mm/dd/yy ") & strSfBeg
               strSfEndDate = Format(DateAdd("d", 1, Now), "mm/dd/yy ") & strSfEnd
            Else
               strSfBegDate = Format(Now, "mm/dd/yy ") & strSfBeg
               strSfEndDate = Format(Now, "mm/dd/yy ") & strSfEnd
            End If

            lSftMin = DateDiff("n", strSfBegDate, strSfEndDate)
            
            lTotMin = lSftMin - lLunMin
            strSfHrs = Format(lTotMin / 60, "##0.00")
            Grd.Col = 4
            Grd.Text = "" & strSfHrs
            
            Grd.Col = 5
            Set Grd.CellPicture = Chkno.Picture
            
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

Private Sub grd_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeySpace Then
      
      Grd.Col = 5
      If (Grd.Rows > 1) Then
         'grd.Col = grd.MouseCol
         If Grd.Row = 0 Then Grd.Row = 1
         If Grd.CellPicture = Chkyes.Picture Then
            Set Grd.CellPicture = Chkno.Picture
         Else
            Set Grd.CellPicture = Chkyes.Picture
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

Private Sub Grd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   
   'grd.Col = grd.MouseCol
   If (Grd.Rows > 1) Then
      Grd.Col = 5
      If Grd.Row = 0 Then Grd.Row = 1
      If Grd.CellPicture = Chkyes.Picture Then
         Set Grd.CellPicture = Chkno.Picture
      Else
         Set Grd.CellPicture = Chkyes.Picture
      End If
   End If
End Sub

Private Sub txtDte_DropDown()
   ShowCalendar Me
End Sub


Private Sub txtDte_LostFocus()
   If (Trim(txtDte) <> "") Then
      GetWeekEnd
      cmdSel.Enabled = True
   End If
End Sub

Private Function GetNewNumber(lEmpID As Long) As String
   Dim S As Single
   Dim l As Long
   Dim m As Long
   Dim t As String
   Dim sNewNumber As String
   
   On Error Resume Next
   '    m = DateValue(Format(ES_SYSDATE, "yyyy,mm,dd"))
   '    s = TimeValue(Format(ES_SYSDATE, "hh:mm:ss"))
   '    l = s * 1000000
   '    GetNewNumber = Format(m, "00000") & Format(l, "000000")
   Dim dt As Variant
   dt = GetServerDateTime()
   m = DateValue(Format(dt, "yyyy,mm,dd"))
   S = TimeValue(Format(dt, "hh:mm:ss"))
   l = S * lEmpID * 10000
   sNewNumber = Format(m, "00000") & Format(l, "000000")
   
   If (Len(sNewNumber) > 11) Then
      GetNewNumber = Mid$(sNewNumber, 1, 11)
   Else
      GetNewNumber = sNewNumber
   End If
   
End Function


Private Sub txtSfBeg_GotFocus()
   SelectFormat Me
End Sub

Private Sub txtSfBeg_LostFocus()
   Dim tc As New ClassTimeCharge
   txtSfBeg = tc.GetTime(txtSfBeg.Text)    'returns blank if invalid
End Sub


Private Sub txtSfEnd_GotFocus()
   SelectFormat Me
End Sub

Private Sub txtSfEnd_LostFocus()
   Dim tc As New ClassTimeCharge
   txtSfEnd = tc.GetTime(txtSfEnd.Text)    'returns blank if invalid
End Sub

Private Sub txtSLBeg_LostFocus()
   Dim tc As New ClassTimeCharge
   txtSLBeg = tc.GetTime(txtSLBeg.Text)    'returns blank if invalid
End Sub
Private Sub txtSLBeg_GotFocus()
   SelectFormat Me
End Sub

Private Sub txtSLEnd_GotFocus()
   SelectFormat Me
End Sub
Private Sub txtSLEnd_LostFocus()
   Dim tc As New ClassTimeCharge
   txtSLEnd = tc.GetTime(txtSLEnd.Text)    'returns blank if invalid
End Sub


Private Function GetTCEntryRecord(strEmp As String, strDate As String) As Boolean
   Dim RdoGet As ADODB.Recordset
   
   On Error GoTo DiaErr1
   sSql = "SELECT * FROM TchdTable where TMEMP = '" & strEmp & "' AND TMDAY = '" & strDate & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoGet, ES_FORWARD)
   
   GetTCEntryRecord = bSqlRows
   
   Set RdoGet = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getweeken"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Function

