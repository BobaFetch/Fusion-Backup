VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form MatlMMf03a 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Set Part QOH to Lot QOH"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8670
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   8670
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox txtDte 
      Height          =   315
      Left            =   1560
      TabIndex        =   20
      Tag             =   "4"
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   5400
      TabIndex        =   19
      Text            =   "Text1"
      Top             =   2160
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton cmdVew 
      DownPicture     =   "MatlMMf03a.frx":0000
      Height          =   350
      Left            =   7080
      Picture         =   "MatlMMf03a.frx":04DA
      Style           =   1  'Graphical
      TabIndex        =   18
      TabStop         =   0   'False
      ToolTipText     =   "Show previous Part QOH"
      Top             =   1800
      Width           =   350
   End
   Begin VB.PictureBox picChkSel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   480
      ScaleHeight     =   150
      ScaleMode       =   0  'User
      ScaleWidth      =   345
      TabIndex        =   6
      Top             =   2280
      Width           =   345
   End
   Begin VB.TextBox txtPtr 
      Height          =   285
      Left            =   1560
      TabIndex        =   4
      Top             =   1800
      Width           =   4215
   End
   Begin VB.CommandButton cmdSel 
      Caption         =   "S&elect"
      Height          =   315
      Left            =   5880
      TabIndex        =   5
      ToolTipText     =   "Select Parts"
      Top             =   1800
      Width           =   1035
   End
   Begin VB.ComboBox cmbCls 
      Height          =   315
      Left            =   3960
      Sorted          =   -1  'True
      TabIndex        =   1
      Tag             =   "3"
      ToolTipText     =   "Select Product Class From List"
      Top             =   480
      Width           =   855
   End
   Begin VB.ComboBox cmbCde 
      Height          =   315
      Left            =   1560
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Select Product Code From List"
      Top             =   480
      Width           =   1215
   End
   Begin VB.ComboBox txtBeg 
      Height          =   315
      Left            =   1560
      TabIndex        =   2
      Tag             =   "4"
      Top             =   840
      Width           =   1335
   End
   Begin VB.ComboBox txtEnd 
      Height          =   315
      Left            =   3720
      TabIndex        =   3
      Tag             =   "4"
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton cmdPQH 
      Caption         =   "&Adjust PA QOH"
      Height          =   435
      Left            =   6840
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "Update Entries and Re-Schedule"
      Top             =   720
      Width           =   1590
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   435
      Left            =   7560
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   120
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   7080
      Top             =   4680
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   5850
      FormDesignWidth =   8670
   End
   Begin MSFlexGridLib.MSFlexGrid grd 
      Height          =   3135
      Left            =   360
      TabIndex        =   10
      ToolTipText     =   "Click The Row To Select A Partnumber to adjust QOH"
      Top             =   2520
      Width           =   7995
      _ExtentX        =   14102
      _ExtentY        =   5530
      _Version        =   393216
      Rows            =   10
      Cols            =   5
      FixedCols       =   0
      BackColorSel    =   -2147483640
      FocusRect       =   0
      HighLight       =   0
      GridLinesFixed  =   1
      ScrollBars      =   2
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      Height          =   255
      Index           =   4
      Left            =   960
      TabIndex        =   21
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label lblSelChk 
      Caption         =   "Check All Part Number"
      Height          =   255
      Left            =   960
      TabIndex        =   17
      Top             =   2280
      Width           =   2655
   End
   Begin VB.Image Chkno 
      Height          =   180
      Left            =   6480
      Picture         =   "MatlMMf03a.frx":09B4
      Stretch         =   -1  'True
      Top             =   5640
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image Chkyes 
      Height          =   180
      Left            =   6000
      Picture         =   "MatlMMf03a.frx":0A0B
      Stretch         =   -1  'True
      Top             =   5640
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Codes"
      Height          =   285
      Index           =   3
      Left            =   240
      TabIndex        =   16
      Top             =   480
      Width           =   1155
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Classes"
      Height          =   285
      Index           =   2
      Left            =   3000
      TabIndex        =   15
      Top             =   480
      Width           =   945
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Through"
      Height          =   255
      Index           =   5
      Left            =   3000
      TabIndex        =   14
      Top             =   840
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include Lots From"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   13
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank for All)"
      Height          =   255
      Index           =   11
      Left            =   5280
      TabIndex        =   12
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank for All)"
      Height          =   255
      Index           =   1
      Left            =   5280
      TabIndex        =   11
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number(s)"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   9
      Top             =   1800
      Width           =   1215
   End
End
Attribute VB_Name = "MatlMMf03a"
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
Dim UsingMouse As Boolean
Dim bQtyDisp As Boolean


Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub cmbCde_LostFocus()
    If Trim(cmbCde) = "" Then cmbCde = "ALL"
End Sub

Private Sub cmbCls_LostFocus()
    If Trim(cmbCls) = "" Then cmbCls = "ALL"
End Sub


Private Sub cmdCan_Click()
   Unload Me
End Sub


Private Sub cmdLHQ_Click()

End Sub

Private Sub cmdPQH_Click()
   Dim bResponse As Byte
   Dim sMsg As String
   
   sMsg = "Do you want to set selected Part's QOH to Lot QOH?"
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      MouseCursor 13
      ' reset the part QOH
      ResetPartQOH
      
      ' Refresh the Grid
      FillPartGrid
      
   End If
   MouseCursor 0
End Sub


Private Sub cmdSel_Click()
    'chkSelAllPrt.value = 0
    FillPartGrid
    MouseCursor 0
End Sub

Private Sub cmdVew_Click()
    MatlMMf03b.Show
End Sub

Private Sub Form_Activate()
    Dim bGoodCal As Boolean
    
    Dim bGoodCoCal As Boolean
    
    If bOnLoad Then
    
      cmbCde.AddItem "ALL"
      FillProductCodes
      cmbCde = "ALL"
      'If Trim(cmbCde) = "" Then cmbCde = cmbCde.List(0)
      
      cmbCls.AddItem "ALL"
      FillProductClasses
      cmbCls = "ALL"
      'If Trim(cmbCls) = "" Then cmbCls = cmbCls.List(0)
  End If
  bOnLoad = 0
  MouseCursor 0
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
      SaveSetting "Esi2000", "EsiAdmn", "AdjustPAQOH", ""
End Sub

Private Sub picChkSel_Click()
    If (picChkSel.Picture = Chkyes.Picture) Then
        picChkSel.Picture = Chkno.Picture
        lblSelChk.Caption = "Check All Part Number"
    Else
        picChkSel.Picture = Chkyes.Picture
        lblSelChk.Caption = "Un-Check All Part Number"
    End If
    
    Dim iList As Integer
    grd.Col = 0
    For iList = 1 To grd.Rows - 1
      ' Only if the part is checked
      grd.Row = iList
      If (picChkSel.Picture = Chkyes.Picture) Then
        Set grd.CellPicture = Chkyes.Picture
      Else
        Set grd.CellPicture = Chkno.Picture
      End If
    Next

End Sub

Private Sub txtBeg_DropDown()
   ShowCalendarEx Me
End Sub

Private Sub txtBeg_LostFocus()
   If Len(Trim(txtBeg)) = 0 Then txtBeg = "ALL"
   If txtBeg <> "ALL" Then txtBeg = CheckDateEx(txtBeg)
End Sub

Private Sub txtEnd_DropDown()
   ShowCalendarEx Me
End Sub

Private Sub txtEnd_LostFocus()
   If Len(Trim(txtEnd)) = 0 Then txtEnd = "ALL"
   If txtEnd <> "ALL" Then txtEnd = CheckDateEx(txtEnd)
End Sub

Private Sub Text1_GotFocus()
   grd.Text = Text1.Text
   If grd.Col >= grd.Cols Then grd.Col = 1
   ChangeCellText
End Sub

Private Sub Grd_EnterCell()  ' Assign cell value to the textbox
   If (bQtyDisp = True) Then Text1.Text = grd.Text
End Sub

Private Sub Grd_LeaveCell()
   ' Assign textbox value to Grd
   If (bQtyDisp = True) And (Text1.Visible = True) Then
      grd.Text = Text1.Text
      Text1.Text = ""
      Text1.Visible = False
   End If

End Sub

Private Sub Text1_LostFocus()

   If (Text1.Visible = True) Then
      grd.Text = Text1.Text
      If (Text1.Text = "") Then
         grd.Col = 0
         Set grd.CellPicture = Chkno.Picture
      End If
      Text1.Text = ""
      Text1.Visible = False
   End If
   
   
   If UsingMouse = True Then
      UsingMouse = False
      Exit Sub
   End If
   
End Sub

Public Sub ChangeCellText() ' Move Textbox to active cell.
   Text1.Move grd.Left + grd.CellLeft, _
   grd.Top + grd.CellTop, _
   grd.CellWidth, grd.CellHeight
   'Text1.SetFocus
   'Text1.ZOrder 0
End Sub

Private Sub Form_Load()
   FormLoad Me
   bOnLoad = 1
   FormatControls
   txtEnd = "ALL"
   txtBeg = "ALL"
    
   With grd
      .ColAlignment(0) = 0
      .ColAlignment(1) = 0
      .ColAlignment(2) = 1
      .ColAlignment(3) = 1
      .ColAlignment(4) = 1
      .Rows = 1
      .Row = 0
      .Col = 0
      .Text = "Sel"
      .Col = 1
      .Text = "Part Number"
      .Col = 2
      .Text = "Part QOH"
      .Col = 3
      .Text = "Lot QOH"
      .Col = 4
      .Text = "Actual QOH"
      .ColWidth(0) = 550
      .ColWidth(1) = 3500
      .ColWidth(2) = 1050
      .ColWidth(3) = 1050
      .ColWidth(4) = 1050
      
   End With
   txtDte = Format(ES_SYSDATE, "mm/dd/yyyy")

   bQtyDisp = False
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   
   FormUnload
   Set MatlMMf03a = Nothing
   
End Sub

Private Sub FillPartGrid()
    Dim RdoGrd As ADODB.Recordset
    Dim strBegDate As String
    Dim strEndDate As String
    Dim strPrt As String
    Dim strCls As String
    Dim strProdCode As String
    Dim strPartNum As String
    
    grd.Rows = 1
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
    
    If (txtPtr.Text = "") Then
        strPrt = "%"
    Else
        strPrt = Compress(txtPtr.Text) & "%"
    End If
    
   If txtBeg = "ALL" Then
      strBegDate = "01/01/1995"
   Else
      strBegDate = Format(txtBeg, "mm/dd/yyyy")
   End If
   
   If txtEnd = "ALL" Then
      strEndDate = "12/31/24"
   Else
      strEndDate = Format(txtEnd, "mm/dd/yyyy")
   End If
    
'    sSql = "SELECT PARTNUM, PAQOH, SUM(LOTREMAININGQTY) LOTREMAININGQTY " & _
'                " FROM PartTable, LohdTable " & _
'            " WHERE LOTPARTREF = PARTREF AND " & _
'                    " LOTPARTREF LIKE '" & strPrt & "' AND " & _
'             " LOTADATE  BETWEEN '" & strBegDate & " 00:00' AND " & _
'                "'" & strEndDate & " 23:59' " & _
'            " GROUP BY PARTNUM,PAQOH " & _
'                "ORDER BY PARTNUM"
'                " AND LOTREMAININGQTY > 0 AND " & _



   sSql = "SELECT PARTNUM,PAQOH,LOTREMAININGQTY,PALOTQTYREMAINING FROM parttable, " & _
            " (select SUM(LOTREMAININGQTY) LOTREMAININGQTY, LOTPARTREF from lohdtable " & _
           "    WHERE   LOTADATE  BETWEEN '" & strBegDate & " 00:00' AND '" & strEndDate & " 23:59' " & _
             "     GROUP BY LOTPARTREF) as f WHERE f.LotPartRef = PartRef  " & _
           " AND PAQOH <> f.LOTREMAININGQTY AND PALEVEL < 5 " & _
           " AND Partref LIKE '" & strPrt & "'"


   bSqlRows = clsADOCon.GetDataSet(sSql, RdoGrd, ES_FORWARD)
   If bSqlRows Then
      With RdoGrd
         Do Until .EOF
            
            If (Trim(!PAQOH) <> Trim(!LOTREMAININGQTY)) Then
            
                grd.Rows = grd.Rows + 1
                grd.Row = grd.Rows - 1
                grd.Col = 0
                Set grd.CellPicture = Chkno.Picture
                
                grd.Col = 1
                grd.Text = "" & Trim(!PartNum)
                grd.Col = 2
                grd.Text = "" & Trim(!PAQOH)
                grd.Col = 3
                grd.Text = "" & Trim(!LOTREMAININGQTY)
                
                Dim strLoitRemQty As String
                strPartNum = Compress(Trim(!PartNum))
                
                'strLoitRemQty = GetLoitRemQty(strPartNum, strBegDate, strEndDate)
                ' update the grid
                'grd.Col = 4
                'grd.Text = "" & strLoitRemQty
                
'                If grd.Rows > 300 Then
'                    MsgBox "There Are more than 300 Parts. Listed first 300 Parts", _
'                       vbInformation, Caption
'                    Exit Do
'                End If
            End If
            
            .MoveNext
         Loop
         ClearResultSet RdoGrd
      End With
      bQtyDisp = True
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

Function GetLoitRemQty(strPartNum As String, strBegDate As String, strEndDate As String)
    Dim RdoLoit As ADODB.Recordset
    Dim strLoitRemQty As String
    On Error GoTo DiaErr1


    sSql = "SELECT SUM(LOIQUANTITY) LOIQUANTITY FROM LoitTable " & _
                " WHERE LOINUMBER IN " & _
            " (SELECT LOTNUMBER FROM LohdTable WHERE " & _
                    " LOTPARTREF = '" & strPartNum & "'" & _
                " AND LOTREMAININGQTY>0 AND " & _
             " LOTADATE  BETWEEN '" & strBegDate & " 00:00' AND " & _
                "'" & strEndDate & " 23:59')"

    bSqlRows = clsADOCon.GetDataSet(sSql, RdoLoit, ES_FORWARD)
    If bSqlRows Then
        With RdoLoit
            strLoitRemQty = IIf(IsNull(!LOIQUANTITY), "-", Trim(!LOIQUANTITY))
        End With
    Else
        strLoitRemQty = "-"
    End If
    Set RdoLoit = Nothing
    
    GetLoitRemQty = strLoitRemQty
    Exit Function
DiaErr1:
   sProcName = "GetLoitRemQty"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Function

Private Sub grd_KeyPress(KeyAscii As Integer)
   On Error Resume Next
   If KeyAscii = vbKeyReturn Or KeyAscii = vbKeySpace Then
      
      If grd.Row >= 1 Then
         If grd.Row = 0 Then grd.Row = 1
            
         If (grd.Col = 0) Then
            grd.Col = 0
            If grd.CellPicture = Chkyes.Picture Then
               Set grd.CellPicture = Chkno.Picture
            Else
               Set grd.CellPicture = Chkyes.Picture
            End If
         ElseIf (grd.Col = 4) Then
            grd.Col = 0
            Set grd.CellPicture = Chkyes.Picture
            grd.Col = 4
            UsingMouse = True
            grd.Text = Text1.Text
            Text1.Visible = True
            ChangeCellText
         End If
            
      End If
   End If
   
End Sub

Private Sub txtDte_DropDown()
   ShowCalendarEx Me
   
End Sub

Private Sub txtDte_LostFocus()
   txtDte = CheckDateEx(txtDte)
   'VerifyDate
   
End Sub


Private Sub grd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   On Error Resume Next
   If grd.Row >= 1 Then
      
      If grd.Row = 0 Then grd.Row = 1
      
      If (grd.Col = 0) Then
         grd.Col = 0
         If grd.CellPicture = Chkyes.Picture Then
            Set grd.CellPicture = Chkno.Picture
         Else
            Set grd.CellPicture = Chkyes.Picture
         End If
      ElseIf (grd.Col = 4) Then
         grd.Col = 0
         Set grd.CellPicture = Chkyes.Picture
         grd.Col = 4
         UsingMouse = True
         grd.Text = Text1.Text
         Text1.Visible = True
         ChangeCellText
      End If
         
   End If
End Sub

Private Function GetCreditAccount(strPartNum As String) As String
   Dim rdoAct As ADODB.Recordset
   
   Dim bType As Byte
   Dim strPCode As String
   
   On Error Resume Next
   'Part First
   sSql = "SELECT PAINVMATACCT, PAPRODCODE FROM PartTable WHERE " _
          & "PARTREF='" & Compress(strPartNum) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoAct, ES_FORWARD)
   If bSqlRows Then
      With rdoAct
         GetCreditAccount = "" & Trim(.Fields(0))
         strPCode = "" & Trim(.Fields(1))
         ClearResultSet rdoAct
      End With
   End If
   If GetCreditAccount = "" Then
      sSql = "SELECT PCINVMATACCT FROM PcodTable WHERE " _
             & "PCREF='" & Compress(strPCode) & "'"
      bSqlRows = clsADOCon.GetDataSet(sSql, rdoAct, ES_FORWARD)
      If bSqlRows Then
         With rdoAct
            GetCreditAccount = "" & Trim(.Fields(0))
            ClearResultSet rdoAct
         End With
      End If
   End If
   
   Set rdoAct = Nothing
   Exit Function
   
DiaErr1:
   'Just bail for now. May not have anything set
   'CurrError.Number = Err
   'CurrError.Description = Err.Description
   'DoModuleErrors Me
   On Error GoTo 0
   
End Function

Private Function GetDebitAccount(strPartNum As String) As String
   Dim rdoAct As ADODB.Recordset
   
   Dim strPCode As String
   Dim strType As String
   
   On Error Resume Next
   GetDebitAccount = ""
   
   'Default Over/Short
   sSql = "SELECT COADJACCT FROM ComnTable WHERE COREF=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoAct, ES_FORWARD)
   If bSqlRows Then
      With rdoAct
         If Not IsNull(!COADJACCT) Then _
                       GetDebitAccount = "" & Trim(!COADJACCT)
         ClearResultSet rdoAct
      End With
   End If
   Set rdoAct = Nothing
   If GetDebitAccount <> "" Then Exit Function
   'Part First
   sSql = "SELECT PACGSMATACCT,PAPRODCODE, PALEVEL  FROM PartTable WHERE " _
          & "PARTREF='" & Compress(strPartNum) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoAct, ES_FORWARD)
   If bSqlRows Then
      With rdoAct
         GetDebitAccount = "" & Trim(.Fields(0))
         strPCode = "" & Trim(.Fields(1))
         strType = "" & Trim(.Fields(2))
         ClearResultSet rdoAct
      End With
   End If
   Set rdoAct = Nothing
   If GetDebitAccount = "" Then
      sSql = "SELECT PCCGSMATACCT FROM PcodTable WHERE " _
             & "PCREF='" & Compress(strPCode) & "'"
      bSqlRows = clsADOCon.GetDataSet(sSql, rdoAct, ES_FORWARD)
      If bSqlRows Then
         With rdoAct
            GetDebitAccount = "" & Trim(.Fields(0))
            ClearResultSet rdoAct
         End With
      End If
      Set rdoAct = Nothing
   End If
   sSql = "SELECT COCGSMATACCT" & Trim(strType) & " " _
          & "FROM ComnTable WHERE COREF=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoAct, ES_FORWARD)
   If bSqlRows Then
      With rdoAct
         If GetDebitAccount = "" Then GetDebitAccount = "" & Trim(.Fields(0))
         ClearResultSet rdoAct
      End With
   End If
   Set rdoAct = Nothing
   Exit Function
   
DiaErr1:
   'Just bail for now. May not have anything set
   'CurrError.Number = Err
   'CurrError.Description = Err.Description
   'DoModuleErrors Me
   On Error GoTo 0
   
End Function

Private Function ResetPartQOH()

   'MouseCursor ccHourglass
   Dim strLotQOH As String
   Dim strPartQOH As String
   Dim strPartNumberRef As String
   Dim strPartNumber As String
   Dim strActLotQOH As String
   Dim strDebitAcct As String
   Dim strCreditAcct As String
   
   Dim iList As Integer
    
   On Error GoTo DiaErr1
   'MouseCursor 13
   Err.Clear
   
   ' Go throught all the record int he grid and re-schedule MO
   For iList = 1 To grd.Rows - 1
      grd.Col = 0
      grd.Row = iList
     
      ' Only if the part is checked
      If grd.CellPicture = Chkyes.Picture Then
        
         grd.Col = 1
         strPartNumberRef = Compress(grd.Text)
         strPartNumber = grd.Text
         
         grd.Col = 2
         strPartQOH = grd.Text
         
         grd.Col = 3
         strLotQOH = grd.Text
         
         grd.Col = 4
         strActLotQOH = grd.Text
         
         If (strActLotQOH <> "") Then
          
            strDebitAcct = GetDebitAccount(strPartNumber)
            strCreditAcct = GetCreditAccount(strPartNumber)
            
            ' Update the PAQOH with the Total Lot QOH
            sSql = "UpdateLotQtyToActualQty '" & strPartNumberRef & "','" & _
                       Format(txtDte, "mm/dd/yyyy") & "'," & strPartQOH & _
                       "," & strLotQOH & "," & strActLotQOH & ",'" & _
                       strDebitAcct & "','" & strCreditAcct & "'"
            
            clsADOCon.ExecuteSQL sSql
            
            If clsADOCon.ADOErrNum <> 0 Then
              MsgBox "Lot Qty For Part Number - " & strPartNumber & " Was Not Adjusted.", _
                 vbInformation, Caption
            End If
                   
          Else
              MsgBox "Actual Lot Qty can not be empty.", _
                 vbInformation, Caption
          
          End If
         ' Update the grid
         grd.Col = 2
         grd.Text = strLotQOH
         
         grd.Col = 0
         Set grd.CellPicture = Chkno.Picture
         
      End If
   Next
   
   MouseCursor 0
   Exit Function

DiaErr1:
   sProcName = "ResetPartQOH"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function


Private Function ResetPartQOH1()

    'MouseCursor ccHourglass
    Dim strLotQOH As String
    Dim strPartQOH As String
    Dim strPartNumberRef As String
    Dim strPartNumber As String
    Dim iList As Integer
    
    On Error GoTo DiaErr1
    'MouseCursor 13
    Err.Clear
   
    ' Go throught all the record int he grid and re-schedule MO
    For iList = 1 To grd.Rows - 1
      grd.Col = 0
      grd.Row = iList
      ' Only if the part is checked
      If grd.CellPicture = Chkyes.Picture Then
        grd.Col = 1
        strPartNumberRef = Compress(grd.Text)
        strPartNumber = grd.Text
        
        
        grd.Col = 2
        strPartQOH = grd.Text
        
        grd.Col = 3
        strLotQOH = grd.Text
        
        ' Update the PAQOH with the Total Lot QOH
        sSql = "UPDATE PartTable SET PAQOH = '" & strLotQOH & "'," & _
                    " PALOTQTYREMAINING = '" & strLotQOH & "'" & _
                " WHERE PARTREF = '" & strPartNumberRef & "'"

        clsADOCon.ExecuteSQL sSql
                
        ' Insert the previous PAQOH in the history table
        sSql = "INSERT INTO MaintPAQOH (PARTREF, PARTNUM, CURPAQOH, " & _
                    " PREPAQOH, PALOTREMAININGQTY )" & _
                "VALUES ('" & strPartNumber & "','" & strPartNumber & "'," & _
                    "'" & strLotQOH & "','" & strPartQOH & _
                    "','" & strLotQOH & "')"
        
        clsADOCon.ExecuteSQL sSql
                    
        ' Update the grid
        grd.Col = 2
        grd.Text = strLotQOH

        grd.Col = 0
        Set grd.CellPicture = Chkno.Picture
        
      End If
    Next
    
    If Err <> 0 Then
       MsgBox "Couldn't Successfully Update..", _
          vbInformation, Caption
    End If

   MouseCursor 0
   Exit Function

DiaErr1:
   sProcName = "ResetPartQOH"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function


