VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form ExportTimeADP 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Export Time Charge to ADP"
   ClientHeight    =   9210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12555
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9210
   ScaleWidth      =   12555
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSelAll 
      Caption         =   "&Select All"
      Height          =   435
      Left            =   10680
      TabIndex        =   18
      TabStop         =   0   'False
      ToolTipText     =   "Create MO from MRP exception"
      Top             =   2040
      Width           =   1755
   End
   Begin VB.CommandButton cmdSel 
      Caption         =   "&Select"
      Height          =   375
      Left            =   3480
      TabIndex        =   17
      TabStop         =   0   'False
      ToolTipText     =   "Enter Updated Time Card"
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox txtFilePath 
      Height          =   285
      Left            =   1800
      TabIndex        =   14
      Tag             =   "3"
      ToolTipText     =   "Select XML file to import"
      Top             =   2040
      Width           =   4695
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "..."
      Height          =   255
      Left            =   6600
      TabIndex        =   13
      ToolTipText     =   "Browse XML file or Text file"
      Top             =   2040
      Width           =   255
   End
   Begin VB.CommandButton cmbExport 
      Caption         =   "Export"
      Height          =   375
      Left            =   7080
      TabIndex        =   12
      Top             =   2040
      Width           =   1335
   End
   Begin VB.TextBox txtPayRate 
      Height          =   285
      Left            =   9420
      TabIndex        =   3
      Tag             =   "2"
      Top             =   480
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.ComboBox cboEnd 
      Height          =   315
      Left            =   1860
      TabIndex        =   2
      Tag             =   "4"
      Top             =   1380
      Width           =   1095
   End
   Begin VB.ComboBox cboEmp 
      Height          =   315
      Left            =   1860
      TabIndex        =   0
      Tag             =   "1"
      ToolTipText     =   "Select From List Or Enter Number"
      Top             =   180
      Width           =   1095
   End
   Begin VB.ComboBox cboStart 
      Height          =   315
      Left            =   1860
      TabIndex        =   1
      Tag             =   "4"
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   7080
      TabIndex        =   4
      Top             =   120
      Width           =   1215
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
      PictureUp       =   "ExportTimeADP.frx":0000
      PictureDn       =   "ExportTimeADP.frx":0146
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5880
      Top             =   1320
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   9210
      FormDesignWidth =   12555
   End
   Begin MSComDlg.CommonDialog fileDlg 
      Left            =   6480
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Open XML File for Import"
      Filter          =   "*.xml"
   End
   Begin MSFlexGridLib.MSFlexGrid Grd 
      Height          =   6375
      Left            =   120
      TabIndex        =   16
      ToolTipText     =   "Click To Select Or Scroll And Press Enter (Also Page Up And Page Down)"
      Top             =   2520
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   11245
      _Version        =   393216
      Rows            =   3
      Cols            =   11
      FixedRows       =   2
      FixedCols       =   0
      FocusRect       =   2
      ScrollBars      =   2
   End
   Begin VB.Image Chkno 
      Height          =   210
      Left            =   8040
      Picture         =   "ExportTimeADP.frx":028C
      Stretch         =   -1  'True
      Top             =   1080
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Chkyes 
      Height          =   210
      Left            =   8040
      Picture         =   "ExportTimeADP.frx":0616
      Stretch         =   -1  'True
      Top             =   840
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Excel File"
      Height          =   285
      Index           =   5
      Left            =   360
      TabIndex        =   15
      Top             =   2040
      Width           =   1305
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Revised Pay Rate"
      Height          =   255
      Index           =   4
      Left            =   7980
      TabIndex        =   11
      Top             =   480
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "End Date"
      Height          =   255
      Index           =   3
      Left            =   420
      TabIndex        =   10
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Number"
      Height          =   255
      Index           =   2
      Left            =   420
      TabIndex        =   9
      Top             =   180
      Width           =   1575
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1860
      TabIndex        =   8
      Top             =   600
      Width           =   2415
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      Height          =   255
      Index           =   0
      Left            =   420
      TabIndex        =   7
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Start Date"
      Height          =   255
      Index           =   1
      Left            =   420
      TabIndex        =   6
      Top             =   1020
      Width           =   975
   End
End
Attribute VB_Name = "ExportTimeADP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of            ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
Option Explicit

Dim bOnLoad As Byte
Dim bGoodCard As Byte
Dim sCardNumber As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cboEmp_Click()
   
    Dim strPremNum As String
    Dim strLastName As String
    Dim strFirstName As String
    Dim iPremNum As Long
    
    Dim RdoEmp As ADODB.Recordset
    
    strPremNum = cboEmp
    
   If (strPremNum = "") Then
      cboEmp = "ALL"
      strPremNum = "ALL"
   End If
   
   If (strPremNum = "ALL") Then
       lblName = " - ALL - "
   Else
       
       sSql = "SELECT PREMLSTNAME, PREMFSTNAME,PREMTERMDT FROM EmplTable WHERE PREMNUMBER = " & CLng(strPremNum)
       bSqlRows = clsADOCon.GetDataSet(sSql, RdoEmp, ES_FORWARD)
       If bSqlRows Then
          With RdoEmp
          
           If (Not IsNull(!PREMTERMDT)) Then
              MsgBox "Not a Current Employee.", vbInformation, Caption
              Set RdoEmp = Nothing
              lblName = "Not a Current Employee"
              Exit Sub
           End If
          
           strLastName = "" & Trim(!PREMLSTNAME)
           strFirstName = "" & Trim(!PREMFSTNAME)
             
           lblName = strLastName & " " & strFirstName
          End With
       End If
       Set RdoEmp = Nothing
   End If
   
End Sub

Private Sub cboEmp_LostFocus()
   cboEmp_Click
End Sub

Private Sub cboEnd_DropDown()
   ShowCalendar Me
End Sub

Private Sub cboEnd_LostFocus()
   cboEnd = CheckDate(cboEnd)
End Sub

Private Sub cmbExport_Click()
   Dim sFileName As String
   
   If (txtFilePath.Text = "") Then
      MsgBox "Please Select Excel File.", vbExclamation
      Exit Sub
   End If
   
   Dim fldCnt As Integer
   Dim sFieldsToExport(12) As String
   
   fldCnt = 11
   AddFieldsToExport sFieldsToExport
   
   sFileName = txtFilePath.Text
   SaveAsExcelGrd sFieldsToExport, fldCnt, sFileName

End Sub

Private Function AddFieldsToExport(ByRef sFieldsToExport() As String)
   
   'Co Code  Batch ID File #   Reg Hours   O/T Hours   Hours 3 Code
   ' Hours 3 Amount Hours 3 Code   Hours 3 Amount Earnings 3 Code
   'Earnings 3 Amount
   Dim I As Integer
   I = 0
   sFieldsToExport(I) = "Co Code"
   sFieldsToExport(I + 1) = "Batch ID"
   sFieldsToExport(I + 2) = "File #"
   sFieldsToExport(I + 3) = "Reg Hours"
   sFieldsToExport(I + 4) = "O/T Hours"
   sFieldsToExport(I + 5) = "Hours 3 Code"
   sFieldsToExport(I + 6) = "Hours 3 Amount"
   sFieldsToExport(I + 7) = "Hours 3 Code"
   sFieldsToExport(I + 8) = "Hours 3 Amount"
   sFieldsToExport(I + 9) = "Earnings 3 Code"
   sFieldsToExport(I + 10) = "Earnings 3 Amount"

End Function

Private Sub cmdSelAll_Click()
   Dim iList As Long
   
   For iList = 1 To Grd.Rows - 1
       Grd.Col = 0
       Grd.Row = iList
       ' Only if the part is checked
       If Grd.CellPicture = Chkno.Picture Then
           Set Grd.CellPicture = Chkyes.Picture
       End If
   Next

End Sub

Private Sub Grd_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
   Dim iCurCol As Integer
   iCurCol = Grd.Col
   If Grd.Row >= 1 Then
      If Grd.Row = 0 Then Grd.Row = 1
      If (Grd.Col = 0) Then
         If Grd.CellPicture = Chkyes.Picture Then
            Set Grd.CellPicture = Chkno.Picture
         Else
            Set Grd.CellPicture = Chkyes.Picture
         End If
      End If
   End If
End Sub

Private Sub grd_KeyPress(KeyAscii As Integer)
   
   If KeyAscii = vbKeyReturn Or KeyAscii = vbKeySpace Then
      Dim iCurCol As Integer
      iCurCol = Grd.Col
      If Grd.Row >= 1 Then
         If Grd.Row = 0 Then Grd.Row = 1
         If (Grd.Col = 0) Then
            If Grd.CellPicture = Chkyes.Picture Then
               Set Grd.CellPicture = Chkno.Picture
            Else
               Set Grd.CellPicture = Chkyes.Picture
            End If
         End If
      End If
   End If

End Sub


Private Sub SaveAsExcelGrd(ByRef aFieldsToExport() As String, _
               iFieldCnt As Integer, ByVal filename)

    Dim xlApp As Excel.Application
    Dim xlBook As Excel.Workbook
    Dim xlSheet As Excel.Worksheet
    Dim iRow, fd, iList As Integer
    Dim iSheetsPerBook As Integer
    
    'Cell count, the cells we can use
    Dim iCell As Integer

    Screen.MousePointer = vbHourglass
    On Error Resume Next
    
    Set xlApp = GetObject(, "Excel.Application")
    On Error GoTo SaveToExcelError
    
    If xlApp Is Nothing Then Set xlApp = New Excel.Application

    iSheetsPerBook = xlApp.SheetsInNewWorkbook
    xlApp.SheetsInNewWorkbook = 1
    Set xlBook = xlApp.Workbooks.Add
    xlApp.SheetsInNewWorkbook = iSheetsPerBook
    xlApp.Visible = False
    xlApp.DisplayAlerts = False
    Set xlSheet = xlBook.Worksheets(1)
    
    'Get the field names
    For fd = 0 To iFieldCnt - 1
      
      xlSheet.Cells(1, fd + 1).Value = aFieldsToExport(fd)
      'xlSheet.Cells(1, fd + 1).Interior.ColorIndex = 33
      'xlSheet.Cells(1, fd + 1).Font.Bold = True
      'xlSheet.Cells(1, fd + 1).BorderAround xlContinuous
    Next
   
   iCell = 0
   iRow = 2
   ' Go throught all the record in the grid and create MO
   For iList = 1 To Grd.Rows - 1
      Grd.Col = 0
      Grd.Row = iList
      ' Only if the part is checked
      If Grd.CellPicture = Chkyes.Picture Then
         
         xlSheet.Cells(iRow, iCell + 1).Value = "XLB"
         
         Grd.Col = 1
         xlSheet.Cells(iRow, iCell + 3).Value = Grd.Text
         
         Grd.Col = 3
         xlSheet.Cells(iRow, iCell + 4).Value = Grd.Text
         
         Grd.Col = 4
         xlSheet.Cells(iRow, iCell + 5).Value = Grd.Text
         
         Grd.Col = 5
         xlSheet.Cells(iRow, iCell + 6).Value = Grd.Text
         
         Grd.Col = 6
         xlSheet.Cells(iRow, iCell + 7).Value = Grd.Text
         
         Grd.Col = 7
         xlSheet.Cells(iRow, iCell + 8).Value = Grd.Text
         Grd.Col = 8
         xlSheet.Cells(iRow, iCell + 9).Value = Grd.Text
         Grd.Col = 9
         xlSheet.Cells(iRow, iCell + 10).Value = Grd.Text
         Grd.Col = 10
         xlSheet.Cells(iRow, iCell + 11).Value = Grd.Text
         
         xlSheet.Columns().AutoFit
         
         iRow = iRow + 1
      End If
      
   Next
   
   xlSheet.Rows.RowHeight = 20

   ' Save the Worksheet.
   If Len(Trim(filename)) > 0 Then
       If InStr(1, filename, ".") = 0 Then filename = filename + ".xlsx"
       xlBook.SaveAs filename
   End If
   
   Set xlSheet = Nothing
   xlBook.Close False
   Set xlBook = Nothing
   xlApp.Visible = True
   xlApp.DisplayAlerts = True
   xlApp.Quit
   Set xlApp = Nothing
   
   
   MsgBox "Successfully Imported The Data."
   Screen.MousePointer = vbArrow
   
   Exit Sub
   
SaveToExcelError:
   MsgBox Err.Description & " Row = " & str(iRow) & " Column = " & str(iCell)


End Sub

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdSearch_Click()
   fileDlg.Filter = "Excel File (*.xls) | *.xls"
   fileDlg.ShowOpen
   If fileDlg.filename = "" Then
       txtFilePath.Text = ""
   Else
       txtFilePath.Text = fileDlg.filename
   End If
End Sub

Private Sub cmdSel_Click()

   Dim iWkNum As Long
   Dim startDt As String
   Dim EndDt As String
   Dim CurWkStartDt As Date
   Dim CurWkEndDt As Date
   
   
    Dim strPremNum As String
    strPremNum = cboEmp
    
   If (strPremNum = "ALL") Then
      strPremNum = ""
   Else
      strPremNum = Val(strPremNum)
   End If
   
   sSql = "DELETE FROM EstmpADPExp"
   clsADOCon.ExecuteSQL sSql
   
   
   startDt = cboStart
   EndDt = cboEnd
   
   CurWkStartDt = CDate(startDt)
   Do While Not (CDate(CurWkStartDt) > CDate(EndDt))
      
      iWkNum = Format(CurWkStartDt, "ww")
      CurWkEndDt = WeekEndDate(iWkNum, 2014)
      If (CDate(CurWkEndDt) > CDate(EndDt)) Then CurWkEndDt = Format(CDate(EndDt), "mm/dd/yyyy")
      ' MM Do some thing
      
      sSql = "INSERT INTO EstmpADPExp (tTCEMP, tWKNUM, tTOTHRS, tOVHRS,tHRSCODE) " & vbCrLf _
            & " SELECT TCEMP, '" & CStr(iWkNum) & "', ROUND(SUM(TCHOURS), 2), " & vbCrLf _
            & " CASE WHEN ROUND(SUM(TCHOURS), 2) - 40 > 0 THEN ROUND(SUM(TCHOURS), 2) - 40 ELSE 0 END, 'RT' " & vbCrLf _
            & " FROM TcitTable join TchdTable on TMCARD = TCCARD " & vbCrLf _
      & " WHERE TMDAY BETWEEN '" & Format(CurWkStartDt, "mm/dd/yyyy") & "'" & vbCrLf _
      & " AND '" & Format(CurWkEndDt, "mm/dd/yyyy") & "' AND TCCODE <> 'VC' " & vbCrLf _
      & " AND TCEMP LIKE '" & strPremNum & "%' GROUP BY TCEMP"
      
      Debug.Print sSql
      
      clsADOCon.ExecuteSQL sSql
      
      sSql = "INSERT INTO EstmpADPExp (tTCEMP, tWKNUM, tTOTHRS, tHRSCODE) " & vbCrLf _
                  & " SELECT TCEMP, '" & CStr(iWkNum) & "', ROUND(SUM(TCHOURS), 2), 'VC' " & vbCrLf _
                  & " FROM TcitTable join TchdTable on TMCARD = TCCARD " & vbCrLf _
               & " WHERE TMDAY BETWEEN '" & Format(CurWkStartDt, "mm/dd/yyyy") & "'" & vbCrLf _
               & " AND '" & Format(CurWkEndDt, "mm/dd/yyyy") & "' AND TCCODE = 'VC' " & vbCrLf _
               & " AND TCEMP LIKE '" & strPremNum & "%' GROUP BY TCEMP"
      
      clsADOCon.ExecuteSQL sSql
      
      CurWkStartDt = Format(CDate(CurWkEndDt) + 1, "mm/dd/yyyy")
      
   Loop
   
   ' Now get the data for the grid.
   FillGrid

End Sub

Function FillGrid() As Integer
    Dim RdoGrd As ADODB.Recordset
    Dim strEmp As String
    
    On Error Resume Next
    Grd.Rows = 1
    On Error GoTo DiaErr1
    
    sSql = "select tTCEMP, PREMLSTNAME,ROUND(SUM(CASE WHEN tHRSCODE = 'RT' THEN tTOTHRS END), 2) as 'TOTHRS'," & _
            " ROUND(SUM(CASE WHEN tHRSCODE = 'RT' THEN tOVHRS END),2) as 'OVHRS'," & _
            " ROUND(SUM(CASE WHEN tHRSCODE = 'VC' THEN tTOTHRS ELSE 0 END), 2) as 'VCHRS'" & _
         " FROM EstmpADPExp, EmplTable WHERE PREMNUMBER = tTCEMP" & _
         " GROUP BY tTCEMP, PREMLSTNAME ORDER BY tTCEMP"

    Debug.Print sSql
    
    bSqlRows = clsADOCon.GetDataSet(sSql, RdoGrd, ES_FORWARD)
    If bSqlRows Then
        With RdoGrd
            Do Until .EOF
            Grd.Rows = Grd.Rows + 1
            Grd.Row = Grd.Rows - 1
            
            Grd.Col = 0
            Set Grd.CellPicture = Chkno.Picture
            Grd.Col = 1
            Grd.Text = "" & Trim(!tTCEMP)
            Grd.Col = 2
            Grd.Text = "" & Trim(!PREMLSTNAME)
            Grd.Col = 3
            Grd.Text = "" & Trim(!TOTHRS)
            Grd.Col = 4
            Grd.Text = "" & Trim(!OVHRS)
            Grd.Col = 5
            Grd.Text = "V"
            Grd.Col = 6
            Grd.Text = "" & Trim(!VCHRS)
            Grd.Col = 7
            Grd.Text = ""
            Grd.Col = 8
            Grd.Text = ""
            Grd.Col = 9
            Grd.Text = ""
            Grd.Col = 10
            Grd.Text = ""
            
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

Private Function WeekEndDate(Week As Long, Optional lngYear As Long) As Date
   Dim WeekStartDate As String
  If lngYear = 0 Then lngYear = Year(Date)
      WeekEndDate = DateSerial(lngYear, 1, 1)
      WeekEndDate = WeekEndDate - Weekday(WeekEndDate) + 7 * (Week)
      
 End Function
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
   
   'fill employee list
   sSql = "select PREMNUMBER from EmplTable where PREMSTATUS <> 'D' " _
          & "order by PREMNUMBER"
   LoadNumComboBox cboEmp, "000000"
   If bSqlRows Then cboEmp = cboEmp.List(0)
   
   cboStart = Format(Now, "mm/dd/yy")
   cboEnd = Format(Now, "mm/dd/yy")
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
      .ColAlignment(8) = 1
      .ColAlignment(9) = 1
      .Rows = 1
      .Row = 0
      .Col = 0
      .Text = "Sel"
      .Col = 1
      .Text = "Emp Number"
      .Col = 2
      .Text = "Last Name"
      .Col = 3
      .Text = "Reg Hrs"
      .Col = 4
      .Text = "O/T Hrs"
      .Col = 5
      .Text = "Hour3 Code"
      .Col = 6
      .Text = "Hour3 Amt"
      .Col = 7
      .Text = "Hour3 Code"
      .Col = 8
      .Text = "Hour3 Amt"
      .Col = 9
      .Text = "Earn3 Code"
      .Col = 10
      .Text = "Earn3 Amt"
      
      .ColWidth(0) = 400
      .ColWidth(1) = 1200
      .ColWidth(2) = 1500
      .ColWidth(3) = 900
      .ColWidth(4) = 900
      .ColWidth(5) = 1000
      .ColWidth(6) = 1000
      .ColWidth(7) = 1000
      .ColWidth(8) = 1000
      .ColWidth(9) = 1200
      .ColWidth(10) = 1200
      
   End With
   
   Show
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set diaHdlch = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub FillCombo()
   On Error GoTo DiaErr1
   sSql = "Qry_FillEmployees"
   LoadNumComboBox cboEmp, "000000"
   If cboEmp.ListCount > 0 Then
      If Trim(sCurrEmployee) = "" Then
         cboEmp = cboEmp.List(0)
      Else
         cboEmp = sCurrEmployee
      End If
      GetEmployee
      cboEmp.AddItem ("ALL")
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub GetEmployee()
   Dim RdoEmp As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "Qry_EmployeeName " & Val(cboEmp)
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoEmp)
   If bSqlRows Then
      With RdoEmp
         cboEmp = Format(!PREMNUMBER, "000000")
         lblName = "" & Trim(!PREMLSTNAME) & ", " _
                   & Trim(!PREMFSTNAME) & " " _
                   & Trim(!PREMMINIT)
         .Cancel
         sCurrEmployee = cboEmp
      End With
   Else
      MsgBox "Employee Wasn't Found.", vbExclamation, Caption
      lblName = "No Current Employee"
   End If
   Set RdoEmp = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getemploy"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub cboStart_DropDown()
   ShowCalendar Me
End Sub


Private Sub cboStart_LostFocus()
   cboStart = CheckDate(cboStart)
End Sub

