VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form MrplMRe01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Create MO's from MRP Exceptions By Part(s)"
   ClientHeight    =   7710
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   11910
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7710
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtPrt 
      Height          =   315
      Left            =   1560
      TabIndex        =   1
      Tag             =   "3"
      Top             =   720
      Width           =   3255
   End
   Begin VB.CommandButton cmdFnd 
      DownPicture     =   "MrplMRe01.frx":0000
      Height          =   315
      Left            =   4920
      Picture         =   "MrplMRe01.frx":0342
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Find A Part Number"
      Top             =   720
      Width           =   350
   End
   Begin VB.CheckBox optSortSchd 
      Alignment       =   1  'Right Justify
      Caption         =   "Sort By Schedule Date"
      Height          =   195
      Left            =   240
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   2520
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   7800
      TabIndex        =   25
      Text            =   "Text1"
      Top             =   2400
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton cmdMO 
      Caption         =   "&Create MO from MRP"
      Height          =   435
      Left            =   9960
      TabIndex        =   24
      TabStop         =   0   'False
      ToolTipText     =   "Create MO from MRP exception"
      Top             =   2880
      Width           =   1755
   End
   Begin VB.PictureBox picUnchecked 
      Height          =   285
      Left            =   8280
      Picture         =   "MrplMRe01.frx":0684
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   23
      Top             =   960
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.PictureBox picChecked 
      Height          =   285
      Left            =   8280
      Picture         =   "MrplMRe01.frx":09C6
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   22
      Top             =   600
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.CommandButton cmdMRP 
      Caption         =   "&Get MRP exception"
      Height          =   435
      Left            =   5640
      TabIndex        =   21
      TabStop         =   0   'False
      ToolTipText     =   "Get MRP exception"
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
      Picture         =   "MrplMRe01.frx":0D08
      Style           =   1  'Graphical
      TabIndex        =   18
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox txtBeg 
      Height          =   315
      Left            =   1560
      TabIndex        =   6
      Tag             =   "4"
      Top             =   1920
      Width           =   1250
   End
   Begin VB.ComboBox txtEnd 
      Height          =   315
      Left            =   4200
      TabIndex        =   7
      Tag             =   "4"
      Top             =   1920
      Width           =   1250
   End
   Begin VB.ComboBox cmbByr 
      Height          =   315
      Left            =   1560
      TabIndex        =   3
      Tag             =   "3"
      Text            =   "ALL"
      ToolTipText     =   "Contains Only Buyers Recorded By The MRP"
      Top             =   1080
      Width           =   2655
   End
   Begin VB.ComboBox cmbCde 
      Height          =   315
      Left            =   1560
      Sorted          =   -1  'True
      TabIndex        =   4
      Tag             =   "3"
      ToolTipText     =   "Select Product Code From List"
      Top             =   1560
      Width           =   1215
   End
   Begin VB.ComboBox cmbCls 
      Height          =   315
      Left            =   4200
      Sorted          =   -1  'True
      TabIndex        =   5
      Tag             =   "3"
      ToolTipText     =   "Select Product Class From List"
      Top             =   1560
      Width           =   855
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6120
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   120
      Width           =   1065
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
      FormDesignHeight=   7710
      FormDesignWidth =   11910
   End
   Begin MSFlexGridLib.MSFlexGrid Grd 
      Height          =   4695
      Left            =   120
      TabIndex        =   20
      ToolTipText     =   "Click To Select Or Scroll And Press Enter (Also Page Up And Page Down)"
      Top             =   2880
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   8281
      _Version        =   393216
      Rows            =   3
      Cols            =   8
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
      Picture         =   "MrplMRe01.frx":14B6
      Stretch         =   -1  'True
      Top             =   720
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Chkyes 
      Height          =   210
      Left            =   6960
      Picture         =   "MrplMRe01.frx":1840
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
      TabIndex        =   19
      Top             =   1080
      Width           =   1155
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank for All)"
      Height          =   255
      Index           =   13
      Left            =   5640
      TabIndex        =   17
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank for All)"
      Height          =   255
      Index           =   11
      Left            =   5640
      TabIndex        =   16
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank for All)"
      Height          =   255
      Index           =   10
      Left            =   5640
      TabIndex        =   15
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Through"
      Height          =   255
      Index           =   9
      Left            =   2880
      TabIndex        =   14
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Start Date"
      Height          =   255
      Index           =   8
      Left            =   240
      TabIndex        =   13
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Buyer"
      Height          =   285
      Index           =   5
      Left            =   240
      TabIndex        =   12
      Top             =   1080
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Classes"
      Height          =   285
      Index           =   0
      Left            =   2880
      TabIndex        =   11
      Top             =   1560
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Codes"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   10
      Top             =   1560
      Width           =   1425
   End
   Begin VB.Label p 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Numbers"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   9
      Top             =   765
      Width           =   1425
   End
End
Attribute VB_Name = "MrplMRe01"
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
Dim bView As Integer

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'Least to greatest dates 10/12/01

Private Sub GetMRPDates()
   Dim RdoDte As ADODB.Recordset
   
    sSql = "SELECT MIN(MRP_PARTDATERQD) FROM MrplTable WHERE " _
           & "MRP_TYPE>" & MRPTYPE_BeginningBalance
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoDte, ES_FORWARD)
   If bSqlRows Then
      With RdoDte
         If Not IsNull(.Fields(0)) Then
            txtBeg = Format(.Fields(0), "mm/dd/yyyy")
         Else
            txtBeg = Format(ES_SYSDATE, "mm/dd/yyyy")
    
         End If
         ClearResultSet RdoDte
      End With
   End If
   txtBeg.ToolTipText = "Earliest Date By Default"
   
    sSql = "SELECT MAX(MRP_PARTDATERQD) FROM MrplTable WHERE " _
           & "MRP_TYPE>" & MRPTYPE_BeginningBalance
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoDte, ES_FORWARD)
   If bSqlRows Then
      With RdoDte
         If Not IsNull(.Fields(0)) Then
            txtEnd = Format(.Fields(0), "mm/dd/yyyy")
         Else
            txtEnd = Format(ES_SYSDATE, "mm/dd/yyyy")
         End If
         ClearResultSet RdoDte
      End With
   End If
   txtEnd.ToolTipText = "Latest Date By Default"
   Set RdoDte = Nothing
End Sub



Private Sub cmbByr_LostFocus()
   cmbByr = CheckLen(cmbByr, 20)
   'If Trim(cmbByr) = "" Then cmbByr = cmbByr.List(0)
   If Trim(cmbByr) = "" Then cmbByr = "ALL"
   
End Sub


Private Sub cmbCde_LostFocus()
   cmbCde = CheckLen(cmbCde, 6)
   If cmbCde = "" Then cmbCde = "ALL"
   
End Sub


Private Sub cmbCls_LostFocus()
   cmbCls = CheckLen(cmbCls, 6)
   If cmbCls = "" Then cmbCls = "ALL"
   
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
   bView = 0
End Sub

Private Sub Text1_GotFocus()
   Grd.Text = Text1.Text
   If Grd.Col >= Grd.Cols Then Grd.Col = 1
   ChangeCellText
End Sub

Private Sub Grd_EnterCell()  ' Assign cell value to the textbox
   If (bGenMRP = True) Then Text1.Text = Grd.Text
End Sub

Private Sub Grd_LeaveCell()
   ' Assign textbox value to Grd
   If (bGenMRP = True) And (Text1.Visible = True) Then
      Grd.Text = Text1.Text
      Text1.Text = ""
      Text1.Visible = False
   End If

End Sub

Private Sub Text1_LostFocus()

   If (Text1.Visible = True) Then
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


Private Sub FillCombos()
    On Error Resume Next
    sSql = "SELECT DISTINCT PARTREF,PARTNUM " _
        & "FROM PartTable  " _
        & "INNER JOIN MrplTable ON MrplTable.MRP_PARTREF=PartTable.PARTREF " _
        & " WHERE PAINACTIVE = 0 AND PAOBSOLETE = 0 " _
        & "ORDER BY PARTREF"
    LoadComboBox cmbPart, 0
    cmbPart = "ALL"
    If Trim(cmbPart) = "" Then cmbPart = "ALL"
End Sub


Private Sub cmdMO_Click()
   CreateNewMO
End Sub

Private Sub cmdMRP_Click()

    Dim sParts As String
    Dim sCode As String
    Dim sClass As String
    Dim sBuyer As String
    Dim sBDate As String
    Dim sEDate As String
    Dim sBegDate As String
    Dim sEndDate As String
   
    Grd.Clear
    GrdAddHeader
    
    GetMRPCreateDates sBegDate, sEndDate
    
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
    If Trim(cmbCde) = "" Then cmbCde = "ALL"
    If Trim(cmbCls) = "" Then cmbCls = "ALL"
    If Trim(cmbByr) = "" Then cmbByr = "ALL"
    If Trim(cmbPart) = "ALL" Then sParts = "" Else sParts = Compress(cmbPart)
    If Trim(cmbCde) = "ALL" Then sCode = "" Else sCode = Compress(cmbCde)
    If Trim(cmbCls) = "ALL" Then sClass = "" Else sClass = Compress(cmbCls)
    If Trim(cmbByr) = "ALL" Then sBuyer = "" Else sBuyer = Trim(cmbByr)
   
   
   Dim RdoMrpEx As ADODB.Recordset
   
   sSql = "SELECT MRP_MOREF, MRP_MONUM, MRP_PARTREF,MRP_PARTNUM," & vbCrLf _
      & "MRP_PARTQTYRQD, CONVERT(varchar(12), MRP_PARTDATERQD,101) MRP_PARTDATERQD," & vbCrLf _
      & "CONVERT(varchar(12), MRP_ACTIONDATE, 101) MRP_actionDt" & vbCrLf _
      & " FROM MrplTable, PartTable " & vbCrLf _
      & " WHERE MRP_PARTREF = PartRef " & vbCrLf _
      & "AND MrplTable.MRP_PARTREF LIKE '" & sParts & "%'" & vbCrLf _
      & "AND MrplTable.MRP_PARTPRODCODE LIKE '" & sCode & "%'" & vbCrLf _
      & "AND MrplTable.MRP_PARTCLASS LIKE '" & sClass & "%'" & vbCrLf _
      & "AND MrplTable.MRP_POBUYER LIKE '" & sBuyer & "%'" & vbCrLf _
      & "AND MrplTable.MRP_PARTDATERQD BETWEEN '" & sBDate & "' AND '" & sEDate & "'" & vbCrLf _
      & "AND MrplTable.MRP_TYPE IN (6, 5) " & vbCrLf _
      & "AND PartTable.PAMAKEBUY ='M'"
   
'not used at this time      & "(select isnull(count(*),0) from BmplTable where BMASSYPART = MRP_PARTREF) as PickCount" & vbCrLf _

   If (optSortSchd.Value = vbChecked) Then
      sSql = sSql & " ORDER BY MrplTable.MRP_ACTIONDATE"
   End If
   
   'Debug.Print sSql
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoMrpEx)
   If bSqlRows Then
      With RdoMrpEx
         Do Until .EOF
            Grd.Rows = Grd.Rows + 1
            Grd.row = Grd.Rows - 1
            
            Grd.Col = 0
            Set Grd.CellPicture = Chkno.Picture
            Grd.Col = 1
            Grd.Text = Trim(!MRP_PARTNUM)
            Grd.Col = 2
            Grd.Text = Trim(!mrp_partqtyrqd)
            Grd.Col = 3
            Grd.Text = Trim(!MRP_PARTDATERQD)
            Grd.Col = 4
            Grd.Text = Trim(!MRP_actionDt)
            Grd.Col = 5
            Set Grd.CellPicture = picUnchecked.Picture
            Grd.Col = 6
            Set Grd.CellPicture = picUnchecked.Picture
            Grd.Col = 7
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
            SelectRunStat iCurCol
            
         ElseIf ((Grd.Col = 5) Or (Grd.Col = 6) Or (Grd.Col = 7)) Then
            
            If Grd.CellPicture = picChecked.Picture Then
               Set Grd.CellPicture = picUnchecked.Picture
            Else
               Set Grd.CellPicture = picChecked.Picture
            End If
            SelectRunStat iCurCol
         ElseIf ((Grd.Col = 2) Or (Grd.Col = 4)) Then
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
         SelectRunStat iCurCol
         
      ElseIf ((Grd.Col = 5) Or (Grd.Col = 6) Or (Grd.Col = 7)) Then
         
         If Grd.CellPicture = picChecked.Picture Then
            Set Grd.CellPicture = picUnchecked.Picture
         Else
            Set Grd.CellPicture = picChecked.Picture
         End If
         SelectRunStat iCurCol
      ElseIf ((Grd.Col = 2) Or (Grd.Col = 3)) Then
         UsingMouse = True
         Grd.Text = Text1.Text
         Text1.Visible = True
         ChangeCellText
   
      End If
   End If
End Sub

Private Sub SelectRunStat(CurCol As Integer)
   
   Dim bPLSel As Boolean
   Dim bSCSel As Boolean
   Dim bMOSel As Boolean
   
   Grd.Col = 0
   bMOSel = IIf((Grd.CellPicture = Chkyes.Picture), True, False)
   
   If (bMOSel = False) And (CurCol = 0) Then
      Grd.Col = 5
      Set Grd.CellPicture = picUnchecked.Picture
      Grd.Col = 6
      Set Grd.CellPicture = picUnchecked.Picture
       
      ' Uncheck both the image
      Exit Sub
   End If
   
   Grd.Col = 5
   bPLSel = IIf((Grd.CellPicture = picChecked.Picture), True, False)
   Grd.Col = 6
   bSCSel = IIf((Grd.CellPicture = picChecked.Picture), True, False)
   
   
   If (bPLSel = False) Then
      Grd.Col = 6
      Set Grd.CellPicture = picChecked.Picture
   Else
      If (CurCol = 6) Then
         Grd.Col = 5
         Set Grd.CellPicture = picUnchecked.Picture
         Grd.Col = 6
         Set Grd.CellPicture = picChecked.Picture
      Else
         Grd.Col = 6
         Set Grd.CellPicture = picUnchecked.Picture
      End If
   End If
   
End Sub

Private Sub Form_Activate()
   On Error Resume Next
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      GetLastMrp
      GetMRPDates
      FillBuyers
      GetOptions
      cmbCde.AddItem "ALL"
      FillProductCodes
      If Trim(cmbCde) = "" Then cmbCde = cmbCde.List(0)
      cmbCls.AddItem "ALL"
      FillProductClasses
      If Trim(cmbCls) = "" Then cmbCls = cmbCls.List(0)
      
      Dim bPartSearch As Boolean
      
      bPartSearch = GetPartSearchOption
      SetPartSearchOption (bPartSearch)
      
      If (Not bPartSearch) Then FillCombos
      
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
      .ColAlignment(2) = 1
      .ColAlignment(3) = 1
      .ColAlignment(4) = 1
      .ColAlignment(5) = 1
      .ColAlignment(6) = 1
      .ColAlignment(7) = 1
   
      .Rows = 1
      .row = 0
      .Col = 0
      .Text = "Sel"
      .Col = 1
      .Text = "PartNumber"
      .Col = 2
      .Text = "Qty"
      .Col = 3
      .Text = "Required Date"
      .Col = 4
      .Text = "Action Date"
      .Col = 5
      .Text = "PL Stat"
      .Col = 6
      .Text = "SC Stat"
      .Col = 7
      .Text = "Print"
      
      .ColWidth(0) = 500
      .ColWidth(1) = 3050
      .ColWidth(2) = 1000
      .ColWidth(3) = 1200
      .ColWidth(4) = 1200
      .ColWidth(5) = 700
      .ColWidth(6) = 700
      .ColWidth(7) = 700
      
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
   Set MrplMRe01 = Nothing
   
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
'   txtPrt = "ALL"
   
End Sub

Private Sub SaveOptions()
   Dim sOptions As String
   Dim sCode As String * 6
   Dim sClass As String * 4
   Dim sBuyer As String * 20
   sCode = cmbCde
   sClass = cmbCls
   sBuyer = cmbByr
   sOptions = sCode & sClass & sBuyer & txtBeg.Text & txtEnd.Text
   SaveSetting "Esi2000", "EsiProd", "MrplMRe01", sOptions
   
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiProd", "MrplMRe01", sOptions)
   If Len(Trim(sOptions)) > 0 Then
      cmbCde = Mid$(sOptions, 1, 6)
      cmbCls = Mid$(sOptions, 7, 4)
      cmbByr = Trim(Mid$(sOptions, 11, 20))
      If Len(sOptions) >= 40 Then txtBeg.Text = Mid$(sOptions, 31, 10)
       If Len(sOptions) >= 50 Then txtEnd.Text = Mid$(sOptions, 41, 10)
   End If
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


Private Sub FillBuyers()
   On Error GoTo DiaErr1
'   sSql = "SELECT DISTINCT MRP_POBUYER FROM MrplTable " _
'          & "WHERE MRP_POBUYER<>'' ORDER BY MRP_POBUYER"
   
   sSql = "SELECT BYREF FROM BuyrTable ORDER BY BYREF"
   
   AddComboStr cmbByr.hwnd, "ALL"
   LoadComboBox cmbByr, -1
   'If Trim(cmbByr) = "" Then cmbByr = cmbByr.List(0)
   If Trim(cmbByr) = "" Then cmbByr = "ALL"
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub cmbPart_LostFocus()
    cmbPart = CheckLen(cmbPart, 30)
    If Trim(cmbPart) = "" Then cmbPart = "ALL"
End Sub

Private Sub txtPrt_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF4 Then
      ViewParts.lblControl = "TXTPRT"
      ViewParts.txtPrt = txtPrt
      ViewParts.Show
   End If
   
End Sub

Private Sub txtPrt_LostFocus()
   txtPrt = CheckLen(txtPrt, 30)
   If Trim(txtPrt) = "" Or Trim(txtPrt) = "ALL" Then txtPrt = "ALL"
   cmbPart = txtPrt
End Sub

Private Sub CreateNewMO()

   Dim iList As Integer
   Dim strPartNum As String
   Dim strQty As String
   Dim strPartRqd As String
   Dim strActDate As String
   Dim strRunStat As String
   Dim strLevel As String
   Dim bPLChked As String
   
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
            strQty = Grd.Text
            Grd.Col = 3
            strPartRqd = Grd.Text
            Grd.Col = 4
            strActDate = Grd.Text
            ' Default va;ue
            strRunStat = "SC"
            bPLChked = False
            Grd.Col = 5
            If (Grd.CellPicture = picChecked.Picture) Then
               strRunStat = "PL"
               bPLChked = True
            End If
            
            Grd.Col = 6
            If (Grd.CellPicture = picChecked.Picture) Then
               strRunStat = "SC"
            End If
            
            Dim strPartRef As String
            Dim iRunNo As Integer
            Dim cPalevLab As Currency
            Dim cPalevExp As Currency
            Dim cPalevMat As Currency
            Dim cPalevOhd As Currency
            Dim cPalevHrs As Currency
            Dim strRouting As String
            
            Dim RdoPart As ADODB.Recordset
            
            sSql = "SELECT PARTNUM, PARTREF, PARUN, PALEVLABOR, PALEVEXP, PALEVMATL," & vbCrLf _
                      & "PALEVOH, PALEVHRS, PALEVEL, PAROUTING " & vbCrLf _
                     & " FROM PartTable WHERE PARTREF = '" & Compress(strPartNum) & "'"

            Debug.Print sSql
            
            bSqlRows = clsADOCon.GetDataSet(sSql, RdoPart)
            If bSqlRows Then
               With RdoPart
                  strPartRef = Trim(!PartRef)
                  iRunNo = CInt(!PARUN) + 1
                  cPalevLab = !PALEVLABOR
                  cPalevExp = !PALEVEXP
                  cPalevMat = !PALEVMATL
                  cPalevOhd = !PALEVOH
                  cPalevHrs = !PALEVHRS
                  strLevel = Trim(!PALEVEL)
                  strRouting = Trim(!PAROUTING)
               End With
            End If
            ClearResultSet RdoPart
            Set RdoPart = Nothing

            ' get Routing information
            Dim RdoRte As ADODB.Recordset
            
            Dim strRoutType As String
            Dim strRtNumber As String
            Dim strRtDesc As String
            Dim strRtBy As String
            Dim strRtAppBy As String
            Dim strRtAppDate As String
            
            sSql = "SELECT * FROM RthdTable WHERE RTREF='" & Compress(strRouting) & "' "
            bSqlRows = clsADOCon.GetDataSet(sSql, RdoRte, ES_FORWARD)
            If bSqlRows Then
               With RdoRte
                  strRtNumber = "" & Trim(!RTNUM)
                  strRtDesc = "" & Trim(!RTDESC)
                  strRtBy = "" & Trim(!RTBY)
                  strRtAppBy = "" & Trim(!RTAPPBY)
                  If Not IsNull(!RTAPPDATE) Then
                     strRtAppDate = Format$(!RTAPPDATE, "mm/dd/yy")
                  Else
                     strRtAppDate = ""
                  End If
                  ClearResultSet RdoRte
               End With
            Else
               strRoutType = "RTEPART" & Trim(strLevel)
               sSql = "SELECT " & strRoutType & " FROM ComnTable WHERE COREF=1"
               Set RdoRte = clsADOCon.GetRecordSet(sSql)
               'Set RdoRte = RdoCon.OpenResultset(sSql, rdOpenForwardOnly, rdConcurReadOnly)
               If Not RdoRte.BOF And Not RdoRte.EOF Then
                  strRtNumber = "" & Trim(RdoRte.Fields(0))
               Else
                  strRtNumber = ""
               End If
               ClearResultSet RdoRte
            End If
            Set RdoRte = Nothing
            
            ' Open the transcation
            clsADOCon.BeginTrans
            clsADOCon.ADOErrNum = 0
            
            ' Create new Runs and schedule
            sSql = "INSERT INTO RunsTable (RUNREF,RUNNO,RUNSCHED," _
               & "RUNSTART, RUNPKSTART, RUNPLDATE," _
               & "RUNSTATUS,RUNQTY,RUNPRIORITY,RUNBUDLAB," _
               & "RUNBUDEXP,RUNBUDMAT,RUNBUDOH,RUNBUDHRS," _
               & "RUNREMAININGQTY,RUNRTNUM,RUNRTDESC,RUNRTBY,RUNRTAPPBY,RUNRTAPPDATE) " _
               & "VALUES('" & strPartRef & "'," _
               & Val(iRunNo) & ",'" _
               & strPartRqd & "','" _
               & strPartRqd & "','" _
               & strPartRqd & "','" _
               & strPartRqd & "','" _
               & strRunStat & "'," _
               & Val(strQty) & "," _
               & Val(0) & "," _
               & cPalevLab & "," _
               & cPalevExp & "," _
               & cPalevMat & "," _
               & cPalevOhd & "," _
               & cPalevHrs & "," _
               & Val(strQty) & ",'" _
               & strRtNumber & "','" _
               & strRtDesc & "','" _
               & strRtBy & "','" _
               & strRtAppBy & "','" _
               & strRtAppDate & "')"
                  
            clsADOCon.ExecuteSql sSql
      
            sSql = "UPDATE PartTable SET PARUN=" & Val(iRunNo) & " " _
                   & "WHERE PARTREF='" & strPartRef & "'"
            clsADOCon.ExecuteSql sSql
            
            ' Now add Routing/Run Op
            CopyRouting Compress(strRtNumber), strPartRef, iRunNo
            
            ' Now add Document list
            CreateDocumentList strPartRef, iRunNo
            
            'now create pick list if requested
            If bPLChked Then
               Dim pk As New ClassPick
               pk.AddPickList strPartRef, iRunNo, CDate(strPartRqd), Val(strQty)
            End If
            
            If clsADOCon.ADOErrNum <> 0 Then
               MsgBox "Couldn't Successfully Update..", _
                  vbInformation, Caption
               clsADOCon.RollbackTrans
            Else
               clsADOCon.CommitTrans
               Grd.Col = 7
'               If ((Grd.CellPicture = picChecked.Picture) And (bPLChked = True)) Then
'                  PrintPickList Trim(strPartRef), Val(iRunNo)
'               End If

               If bPLChked Then
                  PrintMO strPartRef, CStr(iRunNo)
                  If (Grd.CellPicture = picChecked.Picture) Then
                     PrintPickList Trim(strPartRef), Val(iRunNo)
                  End If
               End If
            
               MsgBox "Successfully created MO : " & strPartRef & vbCrLf & _
                     " and Run Number : " & CStr(iRunNo) & ".", _
                           vbInformation, Caption
            End If
        End If
    Next
    
   MouseCursor 0
   Exit Sub

DiaErr1:
   clsADOCon.RollbackTrans
   sProcName = "cmdUpdate"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me

End Sub

Private Function CopyRouting(strRouting As String, strPartRef As String, iRunNo As Integer)
   Dim RdoRte As ADODB.Recordset
   
   Dim iCurrentOp As Integer
   Dim strRoutType As String
   Dim sMsg As String
   
   On Error GoTo DiaErr1
   'Delete possible duplicate keys
   sSql = "DELETE FROM RnopTable WHERE OPREF='" & strPartRef _
          & "' AND OPRUN=" & Val(iRunNo) & " "
   clsADOCon.ExecuteSql sSql
   
   sSql = "SELECT OPREF,OPNO,OPSHOP,OPCENTER,OPSETUP,OPUNIT," _
          & "OPPICKOP,OPSERVPART,OPQHRS,OPMHRS,OPSVCUNIT,OPTOOLLIST,OPCOMT FROM " _
          & "RtopTable WHERE OPREF='" & Compress(strRouting) & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRte, ES_KEYSET)
   If bSqlRows Then
      With RdoRte
         Do Until .EOF
            On Error Resume Next
            If iCurrentOp = 0 Then iCurrentOp = !opNo
            strRoutType = "" & Trim(!OPCOMT)
            strRoutType = ReplaceString(strRoutType)
            sSql = "INSERT INTO RnopTable (OPREF,OPRUN,OPNO,OPSHOP,OPCENTER," _
                   & "OPQHRS,OPMHRS,OPPICKOP,OPSERVPART,OPSUHRS,OPUNITHRS,OPSVCUNIT,OPTOOLLIST,OPCOMT) " _
                   & "VALUES('" & strPartRef & "'," _
                   & Trim(CStr(iRunNo)) & "," _
                   & !opNo & ",'" _
                   & Trim(!OPSHOP) & "','" _
                   & Trim(!OPCENTER) & "'," _
                   & !OPQHRS & "," _
                   & !OPMHRS & "," _
                   & !OPPICKOP & ",'" _
                   & Trim(!OPSERVPART) & "'," _
                   & !OPSETUP & "," _
                   & !OPUNIT & "," _
                   & !OPSVCUNIT & ",'" _
                   & Trim(!OPTOOLLIST) & "','" _
                   & Trim(strRoutType) & "')"
            clsADOCon.ExecuteSql sSql
            .MoveNext
         Loop
         ClearResultSet RdoRte
      End With
      sSql = "UPDATE RunsTable SET RUNOPCUR=" & iCurrentOp & " " _
             & "WHERE RUNREF='" & strPartRef & "' AND RUNNO=" _
             & Val(iRunNo) & " "
      clsADOCon.ExecuteSql sSql
      CopyRouting = 1
   End If
   Set RdoRte = Nothing
   Exit Function
   
DiaErr1:
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   
End Function

Private Sub CreateDocumentList(strPartRef As String, iRunNo As Integer)
   Dim RdoList As ADODB.Recordset
   
   
   Dim iRow As Integer
   Dim sDocRef As String
   Dim sRev As String
   
   On Error GoTo DiaErr1
   sSql = "DELETE FROM RndlTable WHERE RUNDLSRUNREF='" & strPartRef & " ' AND " _
          & "RUNDLSRUNNO=" & Val(iRunNo) & " "
   clsADOCon.ExecuteSql sSql
   
   sSql = "SELECT MAX(DLSREV) FROM DlstTable WHERE DLSREF='" & strPartRef & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoList, ES_KEYSET)
   If bSqlRows Then
      With RdoList
         If Not IsNull(.Fields(0)) Then
            strListRev = "" & Trim(.Fields(0))
         Else
            On Error Resume Next
            'Dummy Row for joins
            sSql = "INSERT INTO RndlTable (RUNDLSNUM,RUNDLSRUNREF, RUNDLSRUNNO) " _
                   & "VALUES(1,'" & strPartRef & "'," & Val(iRunNo) & ")"
            clsADOCon.ExecuteSql sSql
            Exit Sub
         End If
         ClearResultSet RdoList
      End With
   End If
   
   sSql = "DELETE FROM RndlTable WHERE RUNDLSRUNREF='" & strPartRef & " ' AND " _
          & "RUNDLSRUNNO=" & Val(iRunNo) & " "
   clsADOCon.ExecuteSql sSql
   
   ' In partTable the Rev is NONE, but the DocList table has a empty string
   ' 3/7/2010
   If (Trim(strListRev) = "NONE") Then
     strListRev = ""
   End If
   
   sSql = "SELECT * FROM DlstTable WHERE DLSREF='" & strPartRef & "' " _
          & "AND DLSREV='" & strListRev & "' ORDER BY DLSDOCCLASS,DLSDOCREF"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoList, ES_KEYSET)
   If bSqlRows Then
      With RdoList
         On Error Resume Next
         Do Until .EOF
            iRow = iRow + 1
            sDocRef = GetDocInformation("" & Trim(!DLSDOCREF), "" & Trim(!DLSDOCREV))
            sProcName = "CreateDocumentList"
            sSql = "INSERT INTO RndlTable (RUNDLSNUM,RUNDLSRUNREF," _
                   & "RUNDLSRUNNO,RUNDLSREV,RUNDLSDOCREF,RUNDLSDOCREV," _
                   & "RUNDLSDOCREFLONG,RUNDLSDOCREFDESC,RUNDLSDOCREFSHEET," _
                   & "RUNDLSDOCREFCLASS,RUNDLSDOCREFADCN," _
                   & "RUNDLSDOCREFECO) VALUES(" & iRow & ",'" & Compress(strPartRef) & "'," _
                   & Val(iRunNo) & ",'" & strListRev & "','" & Trim(!DLSDOCREF) & "','" _
                   & Trim(!DLSDOCREV) & "','" & strDocName & "','" & strDocDesc & "','" _
                   & strDocSheet & "','" & strDocClass & "','" & strDocAdcn & "'," _
                   & iDocEco & ")"
                   
            clsADOCon.ExecuteSql sSql
            .MoveNext
         Loop
         ClearResultSet RdoList
      End With
      MouseCursor 0
   Else
      'Dummy Row for joins - Corrected 1/30/07
      On Error Resume Next
      sSql = "INSERT INTO RndlTable (RUNDLSNUM,RUNDLSRUNREF, RUNDLSRUNNO) " _
             & "VALUES(1,'" & Compress(strPartRef) & "'," & Val(iRunNo) & ")"
      clsADOCon.ExecuteSql sSql
   End If
   Set RdoList = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getdocumentli"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Function GetDocInformation(DocumentRef As String, DocumentRev As String) As String
   Dim RdoDoc As ADODB.Recordset
   
   sProcName = "getdocinfo"
   sSql = "SELECT DOREF,DONUM,DOREV,DOCLASS,DOSHEET,DODESCR,DOECO," _
          & "DOADCN FROM DdocTable where (DOREF='" & DocumentRef & "' " _
          & "AND DOREV='" & DocumentRev & "')"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoDoc, ES_KEYSET)
   If bSqlRows Then
      With RdoDoc
         GetDocInformation = "" & Trim(!DOREF)
         strDocName = "" & Trim(!DONUM)
         strDocClass = "" & Trim(!DOCLASS)
         strDocSheet = "" & Trim(!DOSHEET)
         strDocDesc = "" & Trim(!DODESCR)
         iDocEco = !DOECO
         strDocAdcn = "" & Trim(!DOADCN)
         ClearResultSet RdoDoc
      End With
      'strDocName = CheckStrings(strDocName)
      'strDocAdcn = CheckStrings(strDocAdcn)
   Else
      strDocName = ""
      strDocClass = ""
      strDocSheet = ""
      strDocDesc = ""
      iDocEco = 0
      strDocAdcn = ""
   End If
   Set RdoDoc = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "GetDocInformation"
   
End Function

Private Sub PrintMO(sPartNumber As String, sRunNo As String)
   MouseCursor 13
   On Error GoTo Psh01
   sProcName = "printreport"
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   Dim sSubSql As String
   
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "PartNumber"
   aFormulaName.Add "RunNumber"
   aFormulaName.Add "ShowOpComments"
   aFormulaName.Add "ShowOpTime"
   aFormulaName.Add "ShowSvcParts"
   aFormulaName.Add "ShowSoAllocs"
   aFormulaName.Add "ShowDocList"
   aFormulaName.Add "ShowBOM"
   aFormulaName.Add "ShowPickList"
   aFormulaName.Add "ShowMoBudget"
   aFormulaName.Add "ShowToolList"
   aFormulaName.Add "ShowServPartDoc"
   aFormulaName.Add "ShowInternalCmt"
   
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'" & CStr(Trim(sPartNumber)) & "'")
   aFormulaValue.Add CStr("'" & sRunNo & "'")
   aFormulaValue.Add CStr("'1'")
   aFormulaValue.Add CStr("'0'")
   aFormulaValue.Add CStr("'0'")
   aFormulaValue.Add CStr("'0'")
   aFormulaValue.Add CStr("'0'")
   aFormulaValue.Add CStr("'0'")
   aFormulaValue.Add CStr("'0'")
   aFormulaValue.Add CStr("'0'")
   aFormulaValue.Add CStr("'0'")
   aFormulaValue.Add CStr("'0'")
   aFormulaValue.Add CStr("'0'")
   
   sCustomReport = GetCustomReport("prdsh01")
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   
   sSql = "{RunsTable.RUNNO} = {@Run} and {PartTable.PARTREF} = {@PartNumber}"
   cCRViewer.SetReportSelectionFormula sSql
   
   sSubSql = "{MopkTable.PKMORUN} = {?Pm-RunsTable.RUNNO} and " _
            & "{MopkTable.PKMORUN} = {?Pm-RunsTable.RUNNO} and " _
            & "{MopkTable.PKMOPART} = {?Pm-RunsTable.RUNREF} and  " _
            & "({MopkTable.PKTYPE} = 10 OR {MopkTable.PKTYPE} = 9)"
            ' PKTYPE=10 is picked type and PickOpenItem = 9
   ' set the sub sql variable pass the sub report name
   cCRViewer.SetSubRptSelFormula "custpklist.rpt", sSubSql

   cCRViewer.CRViewerSize Me
   cCRViewer.SetDbTableConnection
   cCRViewer.ShowGroupTree False
   cCRViewer.OpenCrystalReportObject Me, aFormulaName
      
   cCRViewer.ClearFieldCollection aRptPara
   cCRViewer.ClearFieldCollection aFormulaName
   cCRViewer.ClearFieldCollection aFormulaValue

   MouseCursor 0
   DoEvents
   Exit Sub
   
Psh01:
   sProcName = "printreport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   Resume Psh02
Psh02:
   DoModuleErrors Me
   
End Sub



Private Sub PrintPickList(strPartRef As String, iRunNo As Integer)

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
    aFormulaValue.Add 1
    aFormulaValue.Add 1
    
    cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   
   sSql = "{RunsTable.RUNREF} = '" & strPartRef & "' " _
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

