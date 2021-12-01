VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form PurcPRe07a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Approved Supplier by Part Number"
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9555
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   9555
   ShowInTaskbar   =   0   'False
   Tag             =   "3"
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2760
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   1920
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "PurcPRe07a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdUpd 
      Caption         =   "&Apply"
      Height          =   315
      Left            =   5160
      TabIndex        =   10
      ToolTipText     =   "Update And Apply Changes"
      Top             =   1920
      Width           =   875
   End
   Begin VB.CommandButton cmdSel 
      Caption         =   "&Vendors"
      Height          =   315
      Index           =   1
      Left            =   5160
      TabIndex        =   6
      ToolTipText     =   "Fill The Grid With Vendors (Approved And Not)"
      Top             =   1440
      Width           =   875
   End
   Begin VB.CommandButton cmdSel 
      Caption         =   "&Parts"
      Height          =   315
      Index           =   0
      Left            =   5160
      TabIndex        =   5
      ToolTipText     =   "Select The Parts To Assign"
      Top             =   1080
      Width           =   875
   End
   Begin VB.ComboBox cmbCde 
      Height          =   315
      Left            =   1560
      Sorted          =   -1  'True
      TabIndex        =   4
      Tag             =   "3"
      ToolTipText     =   "Select A Product Code"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   1560
      TabIndex        =   0
      ToolTipText     =   "Contains Up To (300) Part Numbers Equal To Or Greater Than The Entry"
      Top             =   600
      Width           =   3350
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Close"
      Height          =   435
      Left            =   5160
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6600
      Top             =   4200
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   6330
      FormDesignWidth =   9555
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   3855
      Left            =   120
      TabIndex        =   7
      ToolTipText     =   "Click Vendor Or Use SpaceBar To Select - ESC To Cancel"
      Top             =   2400
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   6800
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      Enabled         =   -1  'True
      ScrollBars      =   2
   End
   Begin VB.Label lblEdit 
      BackStyle       =   0  'Transparent
      Caption         =   "Closed (Select Vendors)"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   2040
      Width           =   2400
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Select"
      Height          =   285
      Index           =   3
      Left            =   4560
      TabIndex        =   9
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Select"
      Height          =   285
      Index           =   2
      Left            =   4560
      TabIndex        =   8
      Top             =   1125
      Width           =   1815
   End
   Begin VB.Image Chkno 
      Height          =   210
      Left            =   6120
      Picture         =   "PurcPRe07a.frx":07AE
      Top             =   5520
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Chkyes 
      Height          =   210
      Left            =   6120
      Picture         =   "PurcPRe07a.frx":0B38
      Top             =   5880
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Code"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number(s)"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   1815
   End
End
Attribute VB_Name = "PurcPRe07a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007)) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
'5/14/05 New
Option Explicit
Dim bOnLoad As Byte
Dim bVendors As Integer

Dim sAppVendors() As String
Dim curPrice() As Currency

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbCde_LostFocus()
   If Trim(cmbCde) = "" Then cmbCde = "ALL"
   
End Sub


Private Sub cmbPrt_Change()
   cmdUpd.Enabled = False
   Grid1.Enabled = False
   
End Sub

Private Sub cmbPrt_Click()
   ClearGrid
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub



Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 4308
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub cmdSel_Click(Index As Integer)
   lblEdit = "Closed (Select Vendors)"
   If Index = 0 Then FillParts Else GetApprovals
   
End Sub

Private Sub cmdUpd_Click()
   Dim bResponse As Byte
   bResponse = MsgBox("Update The Current Vendor Appoval Status?", _
               ES_YESQUESTION, Caption)
   If bResponse = vbYes Then UpdateVendors Else CancelTrans
   
End Sub

Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      GetAppVendors
      FillParts
      cmbCde.AddItem "ALL"
      FillProductCodes
      cmbCde = "ALL"
   End If
   bOnLoad = 0
   MouseCursor 0
   
End Sub

Private Sub Text1_GotFocus()
   Grid1.Text = Text1.Text
   If Grid1.Col >= Grid1.Cols Then Grid1.Col = 1
   ChangeCellText
End Sub

Private Sub Grid1_EnterCell()  ' Assign cell value to the textbox
   Text1.Text = Grid1.Text
End Sub

Private Sub Grid1_LeaveCell()
   ' Assign textbox value to Grd
   If Text1.Visible = True Then
      Grid1.Text = Text1.Text
'      If (Trim(Text1.Text) <> "") Then
'         curPrice(Grid1.row, 0) = CCur(Text1.Text)
'      End If
      
      Text1.Text = ""
      Text1.Visible = False
   End If

End Sub

Private Sub Text1_LostFocus()

   If (Text1.Visible = True) Then
      Dim iCurCol As Integer
      Grid1.Text = Text1.Text
'      If (Text1.Text <> "") Then
'         curPrice(Grid1.row, 0) = CCur(Text1.Text)
'      End If
      Text1.Text = ""
      Text1.Visible = False
   End If
   
   
   'If UsingMouse = True Then
   '   UsingMouse = False
   '   Exit Sub
   'End If
   
   
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
   Text1.Move Grid1.Left + Grid1.CellLeft, _
   Grid1.Top + Grid1.CellTop, _
   Grid1.CellWidth, Grid1.CellHeight
   'Text1.SetFocus
   'Text1.ZOrder 0
End Sub


Private Sub Form_Load()
   FormLoad Me
   FormatControls
   sSql = "SELECT PARTREF,PARTNUM FROM PartTable WHERE (PARTREF >= ? " _
          & "AND PAPRODCODE >= ? AND PAINACTIVE = 0 AND PAOBSOLETE = 0)"
   ' Set RdoQry = RdoCon.CreateQuery("", sSql)
   With Grid1
      .Rows = 2
      .ColWidth(0) = 900
      .ColWidth(1) = 1450
      .ColWidth(2) = 2550
      .ColWidth(3) = 850
      .ColWidth(4) = 850
      .ColWidth(5) = 850
      .ColWidth(6) = 850
      
      .ColAlignment(0) = 0
      .ColAlignment(1) = 0
      .ColAlignment(2) = 0
      .row = 0
      .Col = 0
      .Text = "Approved"
      .Col = 1
      .Text = "Vendor "
      .Col = 2
      .Text = "Name"
      .Col = 3
      .Text = "Unit Price"
      .Col = 4
      .Text = "Comment"
      .Col = 5
      .Text = "Unit Price1"
      .Col = 6
      .Text = "Comment1"
      .Col = 0
   End With
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set PurcPRe07a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   cmbCde = "ALL"
   
End Sub

Private Sub FillParts()
   Dim RdoCmb As ADODB.Recordset
   Dim iRows As Integer
   On Error GoTo DiaErr1
   If Trim(cmbCde) <> "ALL" Then
      sSql = "SELECT PARTREF,PARTNUM FROM PartTable WHERE (PARTREF >= '" _
             & Compress(cmbPrt) & "' AND PAPRODCODE ='" & Trim(cmbCde) & "')"
   Else
      sSql = "SELECT PARTREF,PARTNUM FROM PartTable WHERE (PARTREF >= '" _
             & Compress(cmbPrt) & "')"
   End If
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCmb, ES_FORWARD)
   cmbPrt.Clear
   If bSqlRows Then
      With RdoCmb
         Do Until .EOF
            iRows = iRows + 1
            If iRows > 300 Then Exit Do
            AddComboStr cmbPrt.hwnd, "" & Trim(!PartNum)
            .MoveNext
         Loop
         ClearResultSet RdoCmb
      End With
   End If
   If cmbPrt.ListCount > 0 Then
      cmbPrt = cmbPrt.List(0)
      FillVendorGrid
   End If
   Set RdoCmb = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "fillparts"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub FillVendorGrid()
   Dim RdoGrid As ADODB.Recordset
   Dim iRow As Integer
     Grid1.Rows = 1
   
   On Error GoTo DiaErr1
   sSql = "SELECT VEREF,VENICKNAME,VEBNAME FROM VndrTable " _
          & "WHERE VEREF<>'NONE' ORDER BY VEREF"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoGrid, ES_FORWARD)
   If bSqlRows Then
      With RdoGrid
         Do Until .EOF
            iRow = iRow + 1
            Grid1.Rows = iRow + 1
            Grid1.row = iRow
            sAppVendors(iRow, 0) = "" & Trim(!VEREF)
            sAppVendors(iRow, 1) = "0"
            Grid1.Col = 0
            Set Grid1.CellPicture = Chkno.Picture
            Grid1.Col = 1
            Grid1.Text = "" & Trim(!VENICKNAME)
            Grid1.Col = 2
            Grid1.Text = "" & Trim(!VEBNAME)
            'Grid1.Col = 3
            'Grid1.Text = "" & Trim(!AVPARTCOST)
            
            .MoveNext
         Loop
         ClearResultSet RdoGrid
      End With
   End If
   Set RdoGrid = Nothing
   
   If Grid1.Rows > 1 Then Grid1.row = 1
   Exit Sub
   
DiaErr1:
   sProcName = "fillvendorgr"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Function GetAppVendors() As Integer
   On Error Resume Next
   Dim RdoVnd As ADODB.Recordset
   sSql = "SELECT COUNT(VEREF) AS Vendors FROM VndrTable WHERE VEREF<>'NONE'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoVnd, ES_FORWARD)
   If bSqlRows Then
      With RdoVnd
         If Not IsNull(!Vendors) Then GetAppVendors = !Vendors _
                       Else GetAppVendors = 0
         ClearResultSet RdoVnd
      End With
   End If
   Set RdoVnd = Nothing
   
   ReDim sAppVendors(GetAppVendors, 2)
   ReDim curPrice(GetAppVendors, 1)
   
End Function

'Private Sub Grid1_Click()
'   '    MsgBox sAppVendors(Grid1.Row, 0)
'
'End Sub


Private Sub Grid1_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeySpace Then
      Dim iCurCol As Integer
      iCurCol = Grid1.Col
   
      If Grid1.row >= 1 Then
         If (Grid1.Col = 0) Then
            Grid1.Col = 0
            If Grid1.CellPicture = Chkyes.Picture Then
               Set Grid1.CellPicture = Chkno.Picture
               sAppVendors(Grid1.row, 1) = "0"
            Else
               Set Grid1.CellPicture = Chkyes.Picture
               sAppVendors(Grid1.row, 1) = "1"
            End If
         ElseIf ((Grid1.Col = 3)) Then
            Grid1.Col = 0
            Set Grid1.CellPicture = Chkyes.Picture
            sAppVendors(Grid1.row, 1) = "1"
            Grid1.Col = iCurCol

            'UsingMouse = True
            Grid1.Text = Text1.Text
            Text1.Visible = True
            ChangeCellText
         End If
      End If
   End If
   
End Sub


Public Sub UpdateVendors()
   Dim RdoKey As ADODB.Recordset
   Dim bByte As Byte
   Dim iList As Integer
   Dim cost As Currency
   On Error GoTo DiaErr1
   sSql = "DELETE FROM VnapTable WHERE AVPARTREF='" _
          & Compress(cmbPrt) & "'"
   clsADOCon.ExecuteSQL sSql
   sSql = "SELECT AVVENDOR,AVPARTREF,AVPARTCOST, AVCOMMENT, AVPARTCOST1, AVCOMMENT1 FROM VnapTable"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoKey, ES_KEYSET)
   For iList = 1 To Grid1.Rows - 1
      If Val(sAppVendors(iList, 1)) = 1 Then
         With RdoKey
            .AddNew
            Grid1.row = iList
            !AVVENDOR = sAppVendors(iList, 0)
            !AVPARTREF = Compress(cmbPrt)
            
            Grid1.Col = 3
            !AVPARTCOST = Grid1.Text
            Grid1.Col = 4
            !AVCOMMENT = Grid1.Text
            Grid1.Col = 5
            !AVPARTCOST1 = Grid1.Text
            Grid1.Col = 6
            !AVCOMMENT1 = Grid1.Text
            
            '!AVCOMMENT = curPrice(iList, 0)
            '!AVPARTCOST1 = curPrice(iList, 0)
            '!AVCOMMENT1 = curPrice(iList, 0)
            
            .Update
         End With
      End If
   Next
   SysMsg "Part Number Approvals Updated.", True
   Set RdoKey = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "updatevendor"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Public Sub GetApprovals()
   Dim RdoApp As ADODB.Recordset
   Dim iList As Integer
   
   Grid1.Col = 0
   On Error GoTo DiaErr1
   For iList = 1 To Grid1.Rows - 1
      sSql = "SELECT AVVENDOR,AVPARTREF,ISNULL(AVPARTCOST, 0) AVPARTCOST, ISNULL(AVCOMMENT, '') AVCOMMENT, " _
             & " ISNULL(AVPARTCOST1, 0) AVPARTCOST1, ISNULL(AVCOMMENT1, '') AVCOMMENT1 FROM VnapTable " _
             & "WHERE AVVENDOR='" & sAppVendors(iList, 0) & "' " _
             & "AND AVPARTREF='" & Compress(cmbPrt) & "'"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoApp, ES_FORWARD)
      Grid1.Col = 0
      Grid1.row = iList
      If bSqlRows Then
         Set Grid1.CellPicture = Chkyes.Picture
         sAppVendors(Grid1.row, 1) = "1"
         'curPrice(Grid1.row, 0) = RdoApp!AVPARTCOST
         Grid1.Col = 3
         Grid1.Text = RdoApp!AVPARTCOST 'curPrice(Grid1.row, 0)
         Grid1.Col = 4
         Grid1.Text = RdoApp!AVCOMMENT 'curPrice(Grid1.row, 0)
         Grid1.Col = 5
         Grid1.Text = RdoApp!AVPARTCOST1 'curPrice(Grid1.row, 0)
         Grid1.Col = 6
         Grid1.Text = RdoApp!AVCOMMENT1 'curPrice(Grid1.row, 0)
         
      Else
         Set Grid1.CellPicture = Chkno.Picture
         sAppVendors(Grid1.row, 1) = "0"
         Grid1.Col = 3
         Grid1.Text = 0 'curPrice(Grid1.row, 0)
         Grid1.Col = 4
         Grid1.Text = ""
         Grid1.Col = 5
         Grid1.Text = 0
         Grid1.Col = 6
         Grid1.Text = ""
         
      End If
   Next
   Set RdoApp = Nothing
   
   lblEdit = "Edit List"
   cmdUpd.Enabled = True
   Grid1.Enabled = True
   Exit Sub
   
DiaErr1:
   sProcName = "getapprov"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Public Sub ClearGrid()
   Dim iList As Integer
   On Error Resume Next
   lblEdit = "Closed (Select Vendors)"
   Grid1.Col = 0
   For iList = 1 To Grid1.Rows - 1
      Grid1.row = iList
      Set Grid1.CellPicture = Chkno.Picture
      sAppVendors(Grid1.row, 1) = "0"
   Next
   cmdUpd.Enabled = False
   Grid1.Enabled = False
   
End Sub

Private Sub Grid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Grid1.row >= 1 Then
      Dim iCurCol As Integer
      iCurCol = Grid1.Col
      
      If (Grid1.Col = 0) Then
         Grid1.Col = 0
         If Grid1.CellPicture = Chkno.Picture Then
            Set Grid1.CellPicture = Chkyes.Picture
            sAppVendors(Grid1.row, 1) = "1"
         Else
            Set Grid1.CellPicture = Chkno.Picture
            sAppVendors(Grid1.row, 1) = "0"
         End If
      ElseIf ((Grid1.Col = 3) Or (Grid1.Col = 4) Or (Grid1.Col = 5) Or (Grid1.Col = 6)) Then
            Grid1.Col = 0
            Set Grid1.CellPicture = Chkyes.Picture
            sAppVendors(Grid1.row, 1) = "1"
            Grid1.Col = iCurCol
            
            'UsingMouse = True
            Grid1.Text = Text1.Text
            Text1.Visible = True
            ChangeCellText
      End If
   End If

End Sub
