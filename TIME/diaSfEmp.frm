VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaSfEmp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Employee Shift Code"
   ClientHeight    =   4245
   ClientLeft      =   1200
   ClientTop       =   855
   ClientWidth     =   7485
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   7485
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdApy 
      Caption         =   "Apply"
      Height          =   435
      Left            =   5280
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox txtSun 
      Height          =   375
      Left            =   120
      TabIndex        =   28
      Text            =   "Text1"
      Top             =   3600
      Width           =   735
   End
   Begin VB.TextBox txtSat 
      Height          =   375
      Left            =   5760
      TabIndex        =   26
      Text            =   "Text1"
      Top             =   3600
      Width           =   735
   End
   Begin VB.TextBox txtFri 
      Height          =   375
      Left            =   4920
      TabIndex        =   24
      Text            =   "Text1"
      Top             =   3600
      Width           =   735
   End
   Begin VB.TextBox txtThu 
      Height          =   375
      Left            =   3960
      TabIndex        =   22
      Text            =   "Text1"
      Top             =   3600
      Width           =   735
   End
   Begin VB.TextBox txtWed 
      Height          =   375
      Left            =   3000
      TabIndex        =   20
      Text            =   "Text1"
      Top             =   3600
      Width           =   735
   End
   Begin VB.TextBox txtTue 
      Height          =   375
      Left            =   2040
      TabIndex        =   18
      Text            =   "Text1"
      Top             =   3600
      Width           =   735
   End
   Begin VB.TextBox txtMon 
      Height          =   375
      Left            =   1080
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   3600
      Width           =   735
   End
   Begin VB.ComboBox cmbWkSC 
      Height          =   315
      Left            =   5400
      Sorted          =   -1  'True
      TabIndex        =   14
      ToolTipText     =   "Enter/Revise Shift Code (2 char)"
      Top             =   4920
      Width           =   1020
   End
   Begin VB.ComboBox cmbCde 
      Height          =   315
      Left            =   1560
      Sorted          =   -1  'True
      TabIndex        =   2
      ToolTipText     =   "Enter/Revise Shift Code (2 char)"
      Top             =   1800
      Width           =   1020
   End
   Begin VB.Frame z2 
      Height          =   135
      Left            =   120
      TabIndex        =   10
      Top             =   1320
      Width           =   6615
   End
   Begin VB.ComboBox txtEnDte 
      Height          =   315
      Left            =   4200
      TabIndex        =   9
      Tag             =   "4"
      Top             =   5400
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox txtStDte 
      Height          =   315
      Left            =   4080
      TabIndex        =   3
      Tag             =   "4"
      Top             =   4920
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox cmbEmp 
      Height          =   315
      Left            =   1560
      Sorted          =   -1  'True
      TabIndex        =   1
      ToolTipText     =   "Enter/Revise Shift Code"
      Top             =   480
      Width           =   1620
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5640
      Top             =   2640
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4245
      FormDesignWidth =   7485
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Close"
      Height          =   435
      Left            =   5280
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   1230
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   5
      ToolTipText     =   "Subject Help"
      Top             =   0
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   397
      _StockProps     =   65
      BackColor       =   12632256
      Autosize        =   2
      RoundedCorners  =   0   'False
      BevelWidth      =   0
      Outline         =   0   'False
      PictureUp       =   "diaSfEmp.frx":0000
      PictureDn       =   "diaSfEmp.frx":0146
   End
   Begin MSFlexGridLib.MSFlexGrid Grd 
      Height          =   975
      Left            =   240
      TabIndex        =   13
      ToolTipText     =   "Click To Select Or Scroll And Press Enter (Also Page Up And Page Down)"
      Top             =   4800
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   1720
      _Version        =   393216
      Rows            =   3
      Cols            =   7
      FixedRows       =   2
      FixedCols       =   0
      FocusRect       =   2
      ScrollBars      =   2
      MergeCells      =   1
      AllowUserResizing=   3
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Daily Shift Code"
      Height          =   255
      Index           =   11
      Left            =   120
      TabIndex        =   29
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "SUN"
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   27
      Top             =   3360
      Width           =   495
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "SAT"
      Height          =   255
      Index           =   9
      Left            =   5760
      TabIndex        =   25
      Top             =   3360
      Width           =   495
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "FRI"
      Height          =   255
      Index           =   8
      Left            =   4920
      TabIndex        =   23
      Top             =   3360
      Width           =   495
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "THU"
      Height          =   255
      Index           =   7
      Left            =   3960
      TabIndex        =   21
      Top             =   3360
      Width           =   495
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "WED"
      Height          =   255
      Index           =   6
      Left            =   3000
      TabIndex        =   19
      Top             =   3360
      Width           =   495
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "TUE"
      Height          =   255
      Index           =   4
      Left            =   2040
      TabIndex        =   16
      Top             =   3360
      Width           =   495
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "MON"
      Height          =   255
      Index           =   2
      Left            =   1080
      TabIndex        =   15
      Top             =   3360
      Width           =   495
   End
   Begin VB.Label lblShiftDesc 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   1560
      TabIndex        =   12
      Top             =   2280
      Width           =   3615
   End
   Begin VB.Label lblEmpName 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   1560
      TabIndex        =   11
      Top             =   840
      Width           =   3615
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Employee"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   8
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "End Date"
      Height          =   255
      Index           =   5
      Left            =   3120
      TabIndex        =   7
      Top             =   5400
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Start Date"
      Height          =   255
      Index           =   3
      Left            =   3240
      TabIndex        =   6
      Top             =   5040
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Default Shift Code"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   1800
      Width           =   1335
   End
End
Attribute VB_Name = "diaSfEmp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'See the UpdateTables procedure for database revisions
Option Explicit
'Dim RdoCde As ADODB.Recordset
'Dim RdoQry As rdoQuery
Dim AdoCmd As ADODB.Command
Dim AdoParameter As ADODB.Parameter

Dim bOnLoad As Byte
Dim bGoodCode As Byte
Dim sPrevCode As String
Dim bInvalidEmp As Boolean


Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd




Private Function GetEmpShiftDetail() As Byte
    Dim strEmpName As String
    Dim strEmpNumber As String
    
    Dim RdoCde As ADODB.Recordset
    
    strEmpNumber = Compress(cmbEmp) 'GetEmpNumber(strEmpName)
    
    On Error GoTo DiaErr1
    'RdoQry(0) = strEmpNumber
    AdoCmd.Parameters(0).Value = strEmpNumber

    bSqlRows = clsADOCon.GetQuerySet(RdoCde, AdoCmd, ES_FORWARD)
    If bSqlRows Then
        With RdoCde
            cmbCde = "" & Trim(!SFREF)
            GetShiftName (cmbCde.Text)
            lblShiftDesc = GetShiftName(cmbCde.Text)
            FillEmpShiftSch (CStr(cmbEmp))
'            txtStDte = "" & Trim(!startDate)
        End With
        RdoCde.Close
        GetEmpShiftDetail = True
   Else
        cmbCde = ""
'        txtStDte = ""
'        txtEnDte = ""
        GetEmpShiftDetail = False
   End If
   Set RdoCde = Nothing
   
   lblEmpName.Caption = GetEmployeeName(cmbEmp.Text)
   Exit Function
   
DiaErr1:
   sProcName = "GetEmpShiftDetail"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function


Private Function GetEmployeeName(strEmpNo As String) As String
    Dim rdo As ADODB.Recordset
    
    On Error GoTo GetNameError
    GetEmployeeName = ""
    If (strEmpNo <> "") Then
        sSql = "SELECT PREMFSTNAME, PREMLSTNAME FROM EmplTable WHERE PREMNUMBER = " & strEmpNo
        If clsADOCon.GetDataSet(sSql, rdo) Then
            GetEmployeeName = Trim(rdo!PREMLSTNAME & "") & ", " & Trim(rdo!PREMFSTNAME & "")
        End If
    End If
    Set rdo = Nothing
    Exit Function
   
GetNameError:
   sProcName = "getemployeename"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
  
End Function


Private Function GetShiftName(strShiftCode As String) As String
    Dim rdo As ADODB.Recordset
    
    On Error GoTo GetNameError
    GetShiftName = ""
    If (strShiftCode <> "") Then
        sSql = "SELECT SFDESC FROM sfcdTable WHERE SFREF = '" & strShiftCode & "'"
        If clsADOCon.GetDataSet(sSql, rdo) Then GetShiftName = Trim(rdo!SFDESC & "")
    End If
    Set rdo = Nothing
    Exit Function
    
GetNameError:
   sProcName = "getshiftname"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
 
End Function

Private Sub cmbCde_Click()
   lblShiftDesc.Caption = GetShiftName(cmbCde.Text)

   If bOnLoad = 0 Then
      If MsgBox("Update the default shift code to every day?", vbYesNo, "Change Shift Code") = vbYes Then
         FillDialyShift (cmbCde.Text)
      End If
   End If
   
End Sub



Private Sub cmbCde_LostFocus()

   If (bInvalidEmp = False) Then
      If (GetShiftName(cmbCde.Text) <> "") Then
'         If MsgBox("You are changing the shift code for Employee " & lblEmpName & " to " & cmbCde & vbCrLf & "Continue?", vbYesNo, "Change Shift Code") = vbYes Then
'            UpdateShiftCode
'         End If
'         If MsgBox("Update the default shift code to every day?", vbYesNo, "Change Shift Code") = vbYes Then
'            FillDialyShift (cmbCde.Text)
'         End If
         'bGoodCode = GetEmpShiftDetail()
         'FillEmpShiftSch (cmbCde.Text)

      Else
         MsgBox "Not a valid shift code.", vbExclamation, Caption
         lblShiftDesc = ""
         cmbCde.Text = ""
         cmbCde.SetFocus
      End If
   End If
End Sub

Private Sub cmbEmp_Click()
   bInvalidEmp = False
   bGoodCode = GetEmpShiftDetail()
   If Not bGoodCode Then
      AddShiftCode
      FillEmpShiftSch (CStr(cmbEmp))
   End If
End Sub

Private Sub cmbEmp_LostFocus()
   Dim strEmpNumber As String
   ' If not a valid employee then igrnore
   strEmpNumber = Compress(cmbEmp) 'GetEmpNumber(strEmpName)
    
   bInvalidEmp = False
   If (GetEmployeeName(strEmpNumber) <> "") Then
      bGoodCode = GetEmpShiftDetail()
          
      ' If not foudn add the employee
      If Not bGoodCode Then
         AddShiftCode
         FillEmpShiftSch (CStr(cmbEmp))
      End If
   Else
      bInvalidEmp = True
      MsgBox "Not a Valid Employee.", vbExclamation, Caption
      'cmbEmp.TabIndex = 0
      lblEmpName = ""
      cmbEmp.SetFocus
   End If
    
End Sub



Private Sub UpdateShiftCode()
   Dim sShiftCode As String
   sShiftCode = Compress(cmbCde)
   
   clsADOCon.ADOErrNum = 0
   
   On Error GoTo UpdateShiftCodeError1
   sSql = "UPDATE SfempTable SET SFREF='" & sShiftCode & "' WHERE PREMNUMBER=" & cmbEmp.Text
   clsADOCon.ExecuteSQL sSql ' rdExecDirect
      
   'Update the week schedule
   'If MsgBox("Do you apply the shiftcode to all days in a week?", vbYesNo, "Update Weekly shift schedule") = vbYes Then
   UpdateWeeklyShiftCode cmbEmp.Text, sShiftCode
   
    
   On Error Resume Next
   If clsADOCon.ADOErrNum = 0 Then SysMsg "Shift Code Updated", True
   Exit Sub
   
UpdateShiftCodeError1:
    sProcName = "UpdateShiftCode"
    CurrError.Number = clsADOCon.ADOErrNum
    CurrError.Description = Err.Description
    DoModuleErrors Me
End Sub


Private Sub AddShiftCode()
    Dim sShiftCode As String
    If cmbCde.ListCount > 0 Then
        cmbCde.ListIndex = 0
        sShiftCode = Compress(cmbCde)
    Else
        sShiftCode = ""
    End If
    
    On Error GoTo AddShiftCodeError1
        sSql = "INSERT INTO SfempTable (SFREF, PREMNUMBER, STARTDATE, ENDDATE) " _
            & "VALUES ('" & sShiftCode & "', " & cmbEmp.Text & ",'','')"

    clsADOCon.ExecuteSQL sSql ' rdExecDirect
    If clsADOCon.RowsAffected Then
        bGoodCode = GetEmpShiftDetail
        On Error Resume Next
    End If
    Exit Sub
    
AddShiftCodeError1:
    sProcName = "AddShiftCode"
    CurrError.Number = Err.Number
    CurrError.Description = Err.Description
    DoModuleErrors Me
End Sub



Private Sub CmdApy_Click()

   If MsgBox("You are changing the shift code for Employee " & lblEmpName & " to " & cmbCde & vbCrLf & "Continue?", vbYesNo, "Change Shift Code") = vbYes Then
      UpdateShiftCode
   End If

End Sub

Private Sub cmdCan_Click()
    Unload Me
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
    FormLoad Me
    FormatControls
    sSql = "SELECT SFREF, PREMNUMBER,STARTDATE " _
        & " FROM SFEMPTABLE WHERE PREMNUMBER = ?"
   
    'Set RdoQry = RdoCon.CreateQuery("", sSql)
    'RdoQry.MaxRows = 1
   Set AdoCmd = New ADODB.Command
   AdoCmd.CommandText = sSql
    
    Set AdoParameter = New ADODB.Parameter
    AdoParameter.Type = adInteger
    
    AdoCmd.Parameters.Append AdoParameter
        
    bOnLoad = 1
    
     With Grd
      .ColAlignment(0) = 0
      .ColAlignment(1) = 1
      .ColAlignment(2) = 1
      .ColAlignment(3) = 1
      .ColAlignment(4) = 1
      .ColAlignment(5) = 1
      .ColAlignment(6) = 1
   
      .Rows = 1
      .Row = 0
      .Col = 0
      .Text = "SUN"
      .Col = 1
      .Text = "MON"
      .Col = 2
      .Text = "TUE"
      .Col = 3
      .Text = "WED"
      .Col = 4
      .Text = "THU"
      .Col = 5
      .Text = "FRI"
      .Col = 6
      .Text = "SAT"
      
      .ColWidth(0) = 700
      .ColWidth(1) = 700
      .ColWidth(2) = 700
      .ColWidth(3) = 700
      .ColWidth(4) = 700
      .ColWidth(5) = 700
      .ColWidth(6) = 700
      
      .ScrollBars = flexScrollBarBoth
      .AllowUserResizing = flexResizeColumns
      
   End With
    
    Show
End Sub


Private Sub Form_Resize()
   Refresh
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set diaSfEmp = Nothing
End Sub


Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub



Private Sub FillCombo()
    On Error GoTo DiaErr1
    
    sSql = "SELECT DISTINCT PREMLSTNAME, PREMNUMBER FROM EmplTable WHERE PREMTERMDT IS NULL ORDER BY PREMNUMBER"
    LoadComboBox cmbEmp, 0

    If cmbEmp.ListCount > 0 Then
        sSql = "SELECT SFCODE FROM SfcdTable "
        LoadComboBox cmbCde, -1
        LoadComboBox cmbWkSC, -1

        cmbEmp = cmbEmp.List(0)
        bGoodCode = GetEmpShiftDetail
        FillEmpShiftSch (CStr(cmbEmp))
    
    End If

    Exit Sub

DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me

End Sub

Private Function FillDialyShift(strDfShiftCode As String)
   txtSun = strDfShiftCode
   txtMon = strDfShiftCode
   txtTue = strDfShiftCode
   txtWed = strDfShiftCode
   txtThu = strDfShiftCode
   txtFri = strDfShiftCode
   txtSat = strDfShiftCode

End Function

Private Function FillEmpShiftSch(strEmpNo As String)
   On Error GoTo GetNameError
    
   
   Dim RdoSf As ADODB.Recordset
   sSql = "SELECT SFREF, SFREFSUN, SFREFMON, SFREFTUE,SFREFWED," & vbCrLf _
             & "SFREFTHU, SFREFFRI,SFREFSAT " & vbCrLf _
         & " FROM SfempTable " & vbCrLf _
         & " WHERE PREMNUMBER = '" & strEmpNo & "'"
         
   Debug.Print sSql
    
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoSf)
   If bSqlRows Then
      With RdoSf
         
         txtSun = IIf(IsNull(!SFREFSUN), !SFREF, !SFREFSUN)
         txtMon = IIf(IsNull(!SFREFMON), !SFREF, !SFREFMON)
         txtTue = IIf(IsNull(!SFREFTUE), !SFREF, !SFREFTUE)
         txtWed = IIf(IsNull(!SFREFWED), !SFREF, !SFREFWED)
         txtThu = IIf(IsNull(!SFREFTHU), !SFREF, !SFREFTHU)
         txtFri = IIf(IsNull(!SFREFFRI), !SFREF, !SFREFFRI)
         txtSat = IIf(IsNull(!SFREFSAT), !SFREF, !SFREFSAT)
      End With
   End If
   
   Set RdoSf = Nothing
   Exit Function
    
GetNameError:
   sProcName = "FillEmpShiftSch"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me

End Function


Private Function UpdateWeeklyShiftCode(strEmpNo As String, strShiftCode As String)
   
   On Error GoTo GetNameError
   Dim strSFSun As String
   Dim strSFMon As String
   Dim strSFTue As String
   Dim strSFWed As String
   Dim strSFThu As String
   Dim strSFFri As String
   Dim strSFSat As String
   
   strSFSun = IIf(txtSun.Text = "", strShiftCode, txtSun.Text)
   strSFMon = IIf(txtMon.Text = "", strShiftCode, txtMon.Text)
   
   strSFTue = IIf(txtTue.Text = "", strShiftCode, txtTue.Text)
   strSFWed = IIf(txtWed.Text = "", strShiftCode, txtWed.Text)
   strSFThu = IIf(txtThu.Text = "", strShiftCode, txtThu.Text)
   strSFFri = IIf(txtFri.Text = "", strShiftCode, txtFri.Text)
   strSFSat = IIf(txtSat.Text = "", strShiftCode, txtSat.Text)
   
   
   sSql = "UPDATE SfempTable SET SFREFSUN = '" & strSFSun & "',SFREFMON = '" & strSFMon & "'," & vbCrLf _
            & " SFREFTUE ='" & strSFTue & "',SFREFWED = '" & strSFWed & "'," & vbCrLf _
            & " SFREFTHU ='" & strSFThu & "',SFREFFRI = '" & strSFFri & "'," & vbCrLf _
             & " SFREFSAT ='" & strSFSat & "'" & vbCrLf _
         & " WHERE PREMNUMBER = '" & strEmpNo & "'"
         
   clsADOCon.ExecuteSQL sSql ' rdExecDirect
   Exit Function
    
GetNameError:
   sProcName = "FillEmpShiftSch"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me

End Function

