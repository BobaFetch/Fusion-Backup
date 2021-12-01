VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form MrplMRe03 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Revise MO Dates, Quantity and Status"
   ClientHeight    =   9570
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   13020
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9570
   ScaleWidth      =   13020
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdFnd 
      DownPicture     =   "MrplMRe03.frx":0000
      Height          =   315
      Left            =   5160
      Picture         =   "MrplMRe03.frx":0342
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Find A Part Number"
      Top             =   720
      Width           =   350
   End
   Begin VB.TextBox txtPrt 
      Height          =   315
      Left            =   1560
      TabIndex        =   1
      Tag             =   "3"
      Top             =   720
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   7800
      TabIndex        =   20
      Text            =   "Text1"
      Top             =   2400
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton cmdMO 
      Caption         =   "&Update Selected MO"
      Height          =   435
      Left            =   11040
      TabIndex        =   19
      TabStop         =   0   'False
      ToolTipText     =   "Update Selected MO"
      Top             =   2880
      Width           =   1755
   End
   Begin VB.PictureBox picUnchecked 
      Height          =   285
      Left            =   8280
      Picture         =   "MrplMRe03.frx":0684
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   18
      Top             =   960
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.PictureBox picChecked 
      Height          =   285
      Left            =   8280
      Picture         =   "MrplMRe03.frx":09C6
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   17
      Top             =   600
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.CommandButton cmdgetMO 
      Caption         =   "&Get All MO's"
      Height          =   435
      Left            =   5640
      TabIndex        =   16
      TabStop         =   0   'False
      ToolTipText     =   "Get All MO's"
      Top             =   2400
      Width           =   1755
   End
   Begin VB.ComboBox cmbPart 
      Height          =   315
      Left            =   1560
      TabIndex        =   0
      Top             =   720
      Width           =   3495
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "MrplMRe03.frx":0D08
      Style           =   1  'Graphical
      TabIndex        =   13
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox txtBeg 
      Height          =   315
      Left            =   1560
      TabIndex        =   4
      Tag             =   "4"
      Top             =   1560
      Width           =   1095
   End
   Begin VB.ComboBox txtEnd 
      Height          =   315
      Left            =   4200
      TabIndex        =   5
      Tag             =   "4"
      Top             =   1560
      Width           =   1095
   End
   Begin VB.ComboBox cmbRun 
      Height          =   315
      Left            =   1560
      TabIndex        =   3
      Tag             =   "3"
      Text            =   "ALL"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   11400
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   120
      Width           =   1425
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   11280
      Top             =   7320
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   9570
      FormDesignWidth =   13020
   End
   Begin MSFlexGridLib.MSFlexGrid Grd 
      Height          =   6615
      Left            =   120
      TabIndex        =   15
      Top             =   2880
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   11668
      _Version        =   393216
      Rows            =   3
      Cols            =   9
      FixedRows       =   2
      FixedCols       =   0
      RowHeightMin    =   315
      FocusRect       =   2
      ScrollBars      =   2
      MergeCells      =   1
      AllowUserResizing=   3
   End
   Begin VB.Image Chkno 
      Height          =   210
      Left            =   6960
      Picture         =   "MrplMRe03.frx":14B6
      Stretch         =   -1  'True
      Top             =   720
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Chkyes 
      Height          =   210
      Left            =   6960
      Picture         =   "MrplMRe03.frx":1840
      Stretch         =   -1  'True
      Top             =   480
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank for All)"
      Height          =   255
      Index           =   12
      Left            =   5640
      TabIndex        =   14
      Top             =   1080
      Width           =   1155
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank for All)"
      Height          =   255
      Index           =   11
      Left            =   5640
      TabIndex        =   12
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank for All)"
      Height          =   255
      Index           =   10
      Left            =   5640
      TabIndex        =   11
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Through"
      Height          =   255
      Index           =   9
      Left            =   2880
      TabIndex        =   10
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Start Date"
      Height          =   255
      Index           =   8
      Left            =   240
      TabIndex        =   9
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Run"
      Height          =   285
      Index           =   5
      Left            =   240
      TabIndex        =   8
      Top             =   1080
      Width           =   465
   End
   Begin VB.Label p 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Numbers"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   7
      Top             =   765
      Width           =   1425
   End
End
Attribute VB_Name = "MrplMRe03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'5/19/06 Revised report and selections. Removed extra report.
Option Explicit
Dim bOnLoad As Byte

'Passed document stuff
Dim iDocEco As Integer
Dim strDocName As String
Dim strDocClass As String
Dim strDocSheet As String
Dim strDocDesc As String
Dim strDocAdcn As String
Dim sListRef As String
Dim strListRev As String
Dim UsingMouse As Boolean
Dim bGenMRP As Boolean

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'Least to greatest dates 10/12/01

Private Sub cmbRun_LostFocus()
    If Trim(cmbRun) = "" Then cmbRun = "ALL"
End Sub

Private Sub cmdCan_Click()
   Unload Me
   
End Sub


Private Sub cmdFnd_Click()
   If txtPrt.Visible Then
      cmbPart = txtPrt
      ViewParts.lblControl = "TXTPRT"
   Else
      ViewParts.lblControl = "CMBPRT"
   End If
   ViewParts.txtPrt = cmbPart
   ViewParts.Show
End Sub

Private Sub Text1_GotFocus()
   Grd.Text = Text1.Text
   If Grd.Col >= Grd.Cols Then Grd.Col = 1
   ChangeCellText
End Sub

Private Sub Grd_EnterCell()  ' Assign cell value to the textbox
   Text1.Text = Grd.Text
End Sub

Private Sub Grd_LeaveCell()
   ' Assign textbox value to Grd
   If Text1.Visible = True Then
      Grd.Text = Text1.Text
      Text1.Text = ""
      Text1.Visible = False
   End If

End Sub

Private Sub Text1_LostFocus()

   If (Text1.Visible = True) Then
      If (Grd.Col = 4) Then
         If (Not IsDate(Text1.Text)) Then
            MsgBox ("Please enter a valid date.")
            Exit Sub
         End If
      End If
      
      Grd.Text = Text1.Text
      Text1.Text = ""
      Text1.Visible = False
   End If
   
   
   If UsingMouse = True Then
      UsingMouse = False
      Exit Sub
   End If
   
   
'   If Grd.Col <= Grd.Cols - 2 Then
'      Grd.Col = Grd.Col + 1
'      ChangeCellText
'   Else
'      If Grd.row + 1 < Grd.Rows Then
'        Grd.row = Grd.row + 1
'        Grd.Col = 1
'        ChangeCellText
'      End If
'   End If
End Sub

Public Sub ChangeCellText() ' Move Textbox to active cell.
   Text1.Move Grd.Left + Grd.CellLeft, _
   Grd.Top + Grd.CellTop, _
   Grd.CellWidth, Grd.CellHeight
   'Text1.SetFocus
   'Text1.ZOrder 0
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 907
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub


Private Sub FillPartCombos()
   On Error Resume Next
   sSql = "SELECT DISTINCT PARTREF,PARTNUM FROM PartTable,RunsTable " _
        & " WHERE PARTREF=RUNREF AND RUNSTATUS NOT IN('CL', 'CA') " _
       & " ORDER BY PARTREF"
   LoadComboBox cmbPart, 0
   cmbPart = "ALL"
   If Trim(cmbPart) = "" Then cmbPart = "ALL"
   
   FillRun
   
End Sub

Private Sub FillRun()

   On Error Resume Next
   Dim strPart As String
   strPart = Compress(cmbPart)
   cmbRun.Clear
   If (strPart = "ALL") Then
     cmbRun = "ALL"
   Else
      sSql = "select DISTINCT RUNNO from RunsTable WHERE RUNREF = '" & strPart & "' " _
                  & " AND RUNSTATUS NOT IN('CL', 'CA') ORDER BY RUNNO"
      LoadComboBox cmbRun, -1
      
      If Trim(cmbRun) = "" Then cmbRun = "ALL"
   
   End If

End Sub

Private Sub cmdMO_Click()
   CreateNewMO
End Sub

Private Sub cmdgetMO_Click()

    Dim sParts As String
    Dim sRun As String
    Dim sBDate As String
    Dim sEDate As String
    Dim sBegDate As String
    Dim sEndDate As String
   
    Grd.Clear
    GrdAddHeader
    
    If Trim(txtBeg) = "" Then txtBeg = "ALL"
    If Trim(txtEnd) = "" Then txtEnd = "ALL"
    If Not IsDate(txtBeg) Then
       sBDate = "1/1/2000"
    Else
       sBDate = Format(txtBeg, "mm/dd/yyyy")
    End If
    If Not IsDate(txtEnd) Then
       sEDate = "12/31/2024"
    Else
       sEDate = Format(txtEnd, "mm/dd/yyyy")
    End If
    
    If Trim(cmbPart) = "" Then cmbPart = "ALL"
    If Trim(cmbRun) = "" Then cmbRun = "ALL"
    If Trim(cmbPart) = "ALL" Then sParts = "" Else sParts = Compress(cmbPart)
   
   
   Dim RdoMrpEx As ADODB.Recordset
   
   sSql = "select DISTINCT PARTNUM, RUNNO, RUNQTY, RUNSTATUS," & vbCrLf _
         & " CONVERT(varchar(12), RUNSCHED, 101) RUNSCHED, " & vbCrLf _
         & " CONVERT(varchar(12), RUNSTART, 101) RUNSTART " & vbCrLf _
         & " FROM PartTable,RunsTable  WHERE PARTREF = RUNREF AND " & vbCrLf _
         & " RUNSCHED Between '" & sBDate & "' AND '" & sEDate & "'" & vbCrLf _
         & " AND RUNSTATUS NOT LIKE('C%') AND RUNREF LIKE '" & sParts & "%'" & vbCrLf
         
   If (cmbRun <> "ALL") Then sSql = sSql & " AND RUNNO = " & cmbRun
   
   Debug.Print sSql
   Dim strStat As String
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoMrpEx)
   If bSqlRows Then
      With RdoMrpEx
         Do Until .EOF
            Grd.Rows = Grd.Rows + 1
            Grd.row = Grd.Rows - 1
            
            Grd.Col = 0
            Set Grd.CellPicture = Chkno.Picture
            Grd.Col = 1
            Grd.Text = Trim(!PartNum)
            Grd.Col = 2
            Grd.Text = Trim(!Runno)
            Grd.Col = 3
            Grd.Text = Trim(!RUNQTY)
            Grd.Col = 4
            Grd.Text = IIf(Not IsNull(Trim(!RUNSTART)), Trim(!RUNSTART), "")
            Grd.Col = 5
            Grd.Text = IIf(Not IsNull(Trim(!RUNSCHED)), Trim(!RUNSCHED), "")
            Grd.Col = 6
            
            Grd.Text = Trim(!RUNSTATUS)
            strStat = Trim(!RUNSTATUS)
            
            Grd.Col = 7
            If (strStat = "SC" Or strStat = "RL") Then
               Set Grd.CellPicture = picUnchecked.Picture
            Else
               Grd.Text = ""
            End If
            
            Grd.Col = 8
            Set Grd.CellPicture = picUnchecked.Picture
            .MoveNext
         Loop
      End With
   End If
   
   Set RdoMrpEx = Nothing
   bGenMRP = True
   Exit Sub
   
End Sub

Private Sub grd_KeyPress(KeyAscii As Integer)
   
   If KeyAscii = vbKeyReturn Or KeyAscii = vbKeySpace Then
      Dim iCurCol As Integer
      iCurCol = Grd.Col
      If Grd.row >= 1 Then
         If Grd.row = 0 Then Grd.row = 1
               
         If (Grd.Col = 0) Then
            If Grd.CellPicture = Chkyes.Picture Then
               Set Grd.CellPicture = Chkno.Picture
            Else
               Set Grd.CellPicture = Chkyes.Picture
            End If
            Grd.Col = iCurCol
            
         ElseIf (Grd.Col = 7) Then
            
            Grd.Col = 6
            If ((Grd.Text = "RL") Or (Grd.Text = "SC")) Then
               
               Grd.Col = iCurCol
               If Grd.CellPicture = picChecked.Picture Then
                  Set Grd.CellPicture = picUnchecked.Picture
               Else
                  Set Grd.CellPicture = picChecked.Picture
                  Grd.Col = 0
                  Set Grd.CellPicture = Chkyes.Picture
               End If
               
               Grd.Col = iCurCol
            End If
         
         ElseIf (Grd.Col = 8) Then
            If Grd.CellPicture = picChecked.Picture Then
               Set Grd.CellPicture = picUnchecked.Picture
            Else
               Set Grd.CellPicture = picChecked.Picture
               Grd.Col = 0
               Set Grd.CellPicture = Chkyes.Picture
            End If
         ElseIf ((Grd.Col = 3) Or (Grd.Col = 5)) Then
            Grd.Col = 0
            Set Grd.CellPicture = Chkyes.Picture
            Grd.Col = iCurCol
            UsingMouse = True
            Grd.Text = Text1.Text
            Text1.Visible = True
            ChangeCellText
         End If
         
      End If
   End If

End Sub

Private Sub Grd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   
   Dim iCurCol As Integer
   iCurCol = Grd.Col
   If Grd.row >= 1 Then
      If Grd.row = 0 Then Grd.row = 1
            
      If (Grd.Col = 0) Then
         If Grd.CellPicture = Chkyes.Picture Then
            Set Grd.CellPicture = Chkno.Picture
         Else
            Set Grd.CellPicture = Chkyes.Picture
         End If
         Grd.Col = iCurCol
         
      ElseIf (Grd.Col = 7) Then
         
         Grd.Col = 6
         If ((Grd.Text = "RL") Or (Grd.Text = "SC")) Then
            
            Grd.Col = iCurCol
            If Grd.CellPicture = picChecked.Picture Then
               Set Grd.CellPicture = picUnchecked.Picture
            Else
               Set Grd.CellPicture = picChecked.Picture
               Grd.Col = 0
               Set Grd.CellPicture = Chkyes.Picture
            End If
            
            Grd.Col = iCurCol
         End If
      
      ElseIf (Grd.Col = 8) Then
         If Grd.CellPicture = picChecked.Picture Then
            Set Grd.CellPicture = picUnchecked.Picture
         Else
            Set Grd.CellPicture = picChecked.Picture
            Grd.Col = 0
            Set Grd.CellPicture = Chkyes.Picture
         End If
      ElseIf ((Grd.Col = 3) Or (Grd.Col = 5)) Then
         Grd.Col = 0
         Set Grd.CellPicture = Chkyes.Picture
         Grd.Col = iCurCol
         UsingMouse = True
         Grd.Text = Text1.Text
         Text1.Visible = True
         ChangeCellText
      End If
      
   End If
End Sub

Private Sub Form_Activate()
   On Error Resume Next
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      GetOptions
      Dim bPartSearch As Boolean
      
      bPartSearch = GetPartSearchOption
      SetPartSearchOption (bPartSearch)
      
      If (Not bPartSearch) Then FillPartCombos
      
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   GetOptions
   ' Add headers
   GrdAddHeader
   bGenMRP = False
   bOnLoad = 1
   
End Sub

Private Sub GrdAddHeader()
     
     With Grd
      .ColAlignment(0) = 0
      .ColAlignment(1) = 1
      .ColAlignment(2) = 0
      .ColAlignment(3) = 1
      .ColAlignment(4) = 0
      .ColAlignment(5) = 1
      .ColAlignment(6) = 1
      .ColAlignment(7) = 1
      .ColAlignment(8) = 1
   
      .Rows = 1
      .row = 0
      .Col = 0
      .Text = "Sel"
      .Col = 1
      .Text = "PartNumber"
      .Col = 2
      .Text = "Run No"
      .Col = 3
      .Text = "Qty"
      .Col = 4
      .Text = "Start Date"
      .Col = 5
      .Text = "Comp Date"
      .Col = 6
      .Text = "Current Status"
      .Col = 7
      .Text = "PL Stat"
      .Col = 8
      .Text = "Print"
      
      .ColWidth(0) = 500
      .ColWidth(1) = 3050
      .ColWidth(2) = 700
      .ColWidth(3) = 1000
      .ColWidth(4) = 1200
      .ColWidth(5) = 1200
      .ColWidth(6) = 1200
      .ColWidth(7) = 700
      .ColWidth(8) = 700
      
      .ScrollBars = flexScrollBarBoth
      .AllowUserResizing = flexResizeColumns
      
   End With

End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveOptions
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set MrplMRe03 = Nothing
   
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
'   txtPrt = "ALL"
   
End Sub

Private Sub SaveOptions()
   Dim sOptions As String
   Dim sRun As String * 20
   sRun = cmbRun
   SaveSetting "Esi2000", "EsiProd", "MrplMRe03", sOptions
   
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiProd", "MrplMRe03", sOptions)
   If Len(Trim(sOptions)) > 0 Then
      cmbRun = Trim(Mid$(sOptions, 11, 20))
   End If
End Sub

Private Sub txtBeg_DropDown()
   ShowCalendar Me
   
End Sub


Private Sub txtBeg_LostFocus()
   If Len(Trim(txtBeg)) = 0 Then txtBeg = "ALL"
   If txtBeg <> "ALL" Then txtBeg = CheckDate(txtBeg)
   
End Sub


Private Sub txtend_DropDown()
   ShowCalendar Me
   
End Sub


Private Sub txtEnd_LostFocus()
   If Len(Trim(txtEnd)) = 0 Then txtEnd = "ALL"
   If txtEnd <> "ALL" Then txtEnd = CheckDate(txtEnd)
   
End Sub

Private Sub txtPrt_LostFocus()
   txtPrt = CheckLen(txtPrt, 30)
   If Trim(txtPrt) = "" Or Trim(txtPrt) = "ALL" Then txtPrt = "ALL"
   cmbPart = txtPrt
   FillRun
End Sub

Private Sub cmbPart_LostFocus()
   cmbPart = CheckLen(cmbPart, 30)
   If Trim(cmbPart) = "" Then cmbPart = "ALL"
   FillRun
End Sub


Private Sub CreateNewMO()

   Dim iList As Integer
   Dim strPartNum As String
   Dim iRunNo As Integer
   Dim strRunQty As String
   
   Dim strSchDate As String
   Dim strCurRunStat As String
   Dim strNewRunStat As String
   Dim strStartDate As String
   
    On Error GoTo DiaErr1
    MouseCursor 13
    Err.Clear
    
   
    ' Go throught all the record in the grid and create MO
    For iList = 1 To Grd.Rows - 1
        Grd.Col = 0
        Grd.row = iList
        ' Only if the part is checked
        If Grd.CellPicture = Chkyes.Picture Then
            
            Grd.Col = 1
            strPartNum = Grd.Text
            Grd.Col = 2
            iRunNo = Grd.Text
            Grd.Col = 3
            strRunQty = Grd.Text
            Grd.Col = 4
            strStartDate = Grd.Text
            Grd.Col = 5
            strSchDate = Grd.Text
            Grd.Col = 6
            strCurRunStat = Grd.Text
            strNewRunStat = ""
            Grd.Col = 7
            If (Grd.CellPicture = picChecked.Picture) Then
               strNewRunStat = "PL"
            End If
            
            clsADOCon.ADOErrNum = 0
            clsADOCon.BeginTrans
            
            MouseCursor ccHourglass
            
            Dim mo As New ClassMO
            mo.ScheduleOperations Compress(strPartNum), CLng(iRunNo), CCur(strRunQty), CDate(strSchDate), True
            
            MouseCursor ccDefault
            
            If (strNewRunStat <> "") Then
            
               sSql = "UPDATE RunsTable SET RUNSTATUS ='" & strNewRunStat & "' " _
                      & "WHERE RUNREF ='" & Compress(strPartNum) & "' AND " _
                      & " RUNNO = '" & CStr(iRunNo) & "'"
                      
               clsADOCon.ExecuteSQL sSql 'rdExecDirect
            End If
            
            If clsADOCon.ADOErrNum <> 0 Then
               MsgBox "Couldn't Successfully Update..", _
                  vbInformation, Caption
               clsADOCon.RollbackTrans
            Else
               clsADOCon.CommitTrans
               Grd.Col = 8
               If (Grd.CellPicture = picChecked.Picture) Then
                  PrintReport Trim(Compress(strPartNum)), Val(iRunNo)
               End If
            
               MsgBox "Successfully Updated MO : " & strPartNum & vbCrLf & _
                     " and Run Number : " & CStr(iRunNo) & ".", _
                           vbInformation, Caption
            End If
        End If
    Next
    
   MouseCursor 0
   Exit Sub

DiaErr1:
   If clsADOCon.ADOErrNum <> 0 Then
      clsADOCon.RollbackTrans
   End If
   
   sProcName = "cmdUpdate"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me

End Sub



Private Sub PrintReport(strPartRef As String, iRunNo As Integer)

   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection

   MouseCursor 13
   On Error GoTo Pma01Pr

   sCustomReport = GetCustomReport("prdma01")
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport

    aFormulaName.Add "CompanyName"
    aFormulaName.Add "ShowDescription"
    aFormulaName.Add "ShowExtendedDescription"
    aFormulaName.Add "ShowPickComments"
    aFormulaName.Add "ShowLots"
    
    aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
    aFormulaValue.Add 1
    aFormulaValue.Add 0
    aFormulaValue.Add 0
    aFormulaValue.Add 0
    
    cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   
   sSql = "{RunsTable.RUNREF} = '" & Compress(strPartRef) & "' " _
          & "AND {RunsTable.RUNNO}=" & Trim(str(iRunNo)) & " "
   
   cCRViewer.SetReportSelectionFormula sSql
   cCRViewer.SetDbTableConnection
   
   cCRViewer.ShowGroupTree False
   cCRViewer.OpenCrystalReportObject Me, aFormulaName

   cCRViewer.ClearFieldCollection aRptPara
   cCRViewer.ClearFieldCollection aFormulaName
   cCRViewer.ClearFieldCollection aFormulaValue
   
   MouseCursor 0
   Exit Sub
   
Pma01Pr:
   sProcName = "printreport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Function SetPartSearchOption(bPartSearch As Boolean)
   
   If (bPartSearch = True) Then
      cmbPart.Visible = False
      txtPrt.Visible = True
      cmdFnd.Visible = True
   Else
      cmbPart.Visible = True
      txtPrt.Visible = False
      cmdFnd.Visible = False
   End If
End Function



