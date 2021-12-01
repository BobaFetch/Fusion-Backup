VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form LotsLTe02a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Set Lot Tracking Requirements"
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6825
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H8000000F&
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   6825
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   1920
      TabIndex        =   0
      Top             =   480
      Width           =   3075
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "LotsLTe02a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   28
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.Frame z2 
      Height          =   60
      Left            =   360
      TabIndex        =   26
      Top             =   1680
      Width           =   6405
   End
   Begin VB.CommandButton cmdUpd 
      Caption         =   "&Apply"
      Enabled         =   0   'False
      Height          =   315
      Left            =   5880
      TabIndex        =   25
      ToolTipText     =   "Update The List Of Changes"
      Top             =   1800
      Width           =   875
   End
   Begin VB.CommandButton cmdNil 
      Cancel          =   -1  'True
      Caption         =   "N&one"
      Enabled         =   0   'False
      Height          =   315
      Left            =   2760
      TabIndex        =   13
      ToolTipText     =   "Mark All Parts In Grid As Not Requiring On Dock Inspection"
      Top             =   1800
      Width           =   800
   End
   Begin VB.CommandButton cmdAll 
      Caption         =   "&All"
      Enabled         =   0   'False
      Height          =   315
      Left            =   1920
      TabIndex        =   12
      ToolTipText     =   "Mark All Parts In Grid As Requiring On Dock Inspection"
      Top             =   1800
      Width           =   800
   End
   Begin VB.CheckBox typ 
      Caption         =   "1"
      Height          =   255
      Index           =   1
      Left            =   1920
      TabIndex        =   2
      Top             =   840
      Value           =   1  'Checked
      Width           =   495
   End
   Begin VB.CheckBox typ 
      Caption         =   "2"
      Height          =   255
      Index           =   2
      Left            =   2400
      TabIndex        =   3
      Top             =   840
      Value           =   1  'Checked
      Width           =   495
   End
   Begin VB.CheckBox typ 
      Caption         =   "3"
      Height          =   255
      Index           =   3
      Left            =   2880
      TabIndex        =   4
      Top             =   840
      Value           =   1  'Checked
      Width           =   495
   End
   Begin VB.CheckBox typ 
      Caption         =   "4"
      Height          =   255
      Index           =   4
      Left            =   3360
      TabIndex        =   5
      Top             =   840
      Value           =   1  'Checked
      Width           =   495
   End
   Begin VB.CheckBox typ 
      Caption         =   "5"
      Height          =   255
      Index           =   5
      Left            =   3840
      TabIndex        =   6
      Top             =   840
      Value           =   1  'Checked
      Width           =   495
   End
   Begin VB.CheckBox typ 
      Caption         =   "6"
      Height          =   255
      Index           =   6
      Left            =   4320
      TabIndex        =   7
      Top             =   840
      Value           =   1  'Checked
      Width           =   495
   End
   Begin VB.CheckBox typ 
      Caption         =   "7"
      Height          =   255
      Index           =   7
      Left            =   4800
      TabIndex        =   8
      Top             =   840
      Value           =   1  'Checked
      Width           =   495
   End
   Begin VB.CheckBox typ 
      Caption         =   "8"
      Height          =   255
      Index           =   8
      Left            =   5280
      TabIndex        =   9
      Top             =   840
      Value           =   1  'Checked
      Width           =   495
   End
   Begin VB.CheckBox optVew 
      Caption         =   "View"
      Height          =   255
      Left            =   600
      TabIndex        =   22
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtPrt 
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Tag             =   "3"
      Text            =   "ALL"
      ToolTipText     =   "Part Number(Blank For All) Equal To Or Greater Than Selection"
      Top             =   510
      Visible         =   0   'False
      Width           =   3075
   End
   Begin VB.CommandButton cmdFnd 
      DownPicture     =   "LotsLTe02a.frx":07AE
      Height          =   315
      Left            =   5040
      Picture         =   "LotsLTe02a.frx":0AF0
      Style           =   1  'Graphical
      TabIndex        =   21
      TabStop         =   0   'False
      ToolTipText     =   "Find A Part Number"
      Top             =   480
      Visible         =   0   'False
      Width           =   350
   End
   Begin VB.CommandButton cmdSel 
      Caption         =   "S&elect"
      Height          =   315
      Left            =   3840
      TabIndex        =   11
      ToolTipText     =   "Fill The Grid (300 Parts Maximum)"
      Top             =   1200
      Width           =   875
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   3135
      Left            =   360
      TabIndex        =   14
      ToolTipText     =   "Double Click Or Press Enter To Add Or Remove Entry"
      Top             =   2160
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   5530
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      Enabled         =   0   'False
      ScrollBars      =   2
   End
   Begin VB.ComboBox cmbCde 
      Height          =   315
      Left            =   1920
      TabIndex        =   10
      Tag             =   "3"
      ToolTipText     =   "Select Product Code Or Leave Blank For All)"
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Close"
      Height          =   435
      Left            =   5880
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   8760
      Top             =   3360
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   5790
      FormDesignWidth =   6825
   End
   Begin MSComctlLib.ProgressBar prg1 
      Height          =   300
      Left            =   2160
      TabIndex        =   29
      Top             =   5400
      Width           =   4092
      _ExtentX        =   7223
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Image Chkno 
      Height          =   210
      Left            =   6360
      Picture         =   "LotsLTe02a.frx":0E32
      Top             =   600
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Chkyes 
      Height          =   210
      Left            =   6120
      Picture         =   "LotsLTe02a.frx":11BC
      Top             =   600
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Label lblInActive 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Lot Tracking Has Not Been Activated.  No Changes Will Be Saved"
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1800
      TabIndex        =   27
      Top             =   0
      Width           =   3975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Mark Part Numbers:"
      Height          =   255
      Index           =   4
      Left            =   360
      TabIndex        =   24
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Types?"
      Height          =   285
      Index           =   1
      Left            =   360
      TabIndex        =   23
      Top             =   840
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Code(s)"
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   20
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label lblInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   5400
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   285
      Left            =   6120
      TabIndex        =   18
      ToolTipText     =   "Fills A Maximun Of 300 Part Numbers"
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Selected"
      Height          =   255
      Index           =   2
      Left            =   4920
      TabIndex        =   17
      ToolTipText     =   "Fills A Maximun Of 300 Part Numbers"
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number(s)"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   16
      Top             =   480
      Width           =   1455
   End
End
Attribute VB_Name = "LotsLTe02a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Stanwood, Washington, USA  ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'New 1/14/03
'9/1/04 omit tools
'4/28/06 Aligned Check left and cleaned up query
Option Explicit

Dim bOnLoad As Byte
Dim bRefreshed As Byte

Dim lSonumber As Long
Dim iItem As Integer
Dim sRev As String
Dim sDesc As String

'Varibles for recording change
Dim bLtChng(300) As Byte
Dim bLtOrig(300) As Byte
Dim bLtReqd(300) As Byte
Dim sLtPart(300) As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd



Private Sub cmbCde_LostFocus()
   cmbCde = CheckLen(cmbCde, 6)
   If cmbCde = "" Then cmbCde = "ALL"
   
End Sub


Private Sub cmdAll_Click()
   Dim iList As Integer
   If lblInActive.Visible = False Then cmdUpd.Enabled = True
   For iList = 1 To Grid1.Rows - 1
      Grid1.Row = iList
      Grid1.Col = 0
      Set Grid1.CellPicture = Chkyes.Picture
      bLtReqd(Grid1.Row) = 1
      bLtChng(Grid1.Row) = 1
   Next
   
End Sub

Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdFnd_Click()
   ViewParts.lblControl = "TXTPRT"
   ViewParts.txtPrt = txtPrt
   optVew.Value = vbChecked
   ViewParts.Show
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext "5502"
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub


Private Sub cmdNil_Click()
   Dim iList As Integer
   On Error Resume Next
   If lblInActive.Visible = False Then cmdUpd.Enabled = True
   For iList = 1 To Grid1.Rows - 1
      Grid1.Row = iList
      Grid1.Col = 0
      Set Grid1.CellPicture = Chkno.Picture
      bLtReqd(Grid1.Row) = 0
      bLtChng(Grid1.Row) = 1
   Next
   
End Sub

Private Sub cmdSel_Click()
   FillGrid
   
End Sub

Private Sub cmdUpd_Click()
   Dim bResponse As Byte
   Dim sMsg As String
   
   sMsg = "You Have Chosen To Update Changed Part Number Status" & vbCr _
          & "Of Lot Tracking.  Do You Wish To Continue?"
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      UpdateList
   Else
      CancelTrans
   End If
   
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   cmbCde.AddItem "ALL"
   cmbCde = cmbCde.List(0)
   If bOnLoad Then
      FillProductCodes
      FillCombo
      cmbPrt = "ALL"
      'Lots working?
      bOnLoad = CheckLotStatus()
      If bOnLoad = 1 Then lblInActive.Visible = False
   End If
   bOnLoad = 0
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   Height = 5965
   FormLoad Me
   FormatControls
   
   GetOptions
   With Grid1
      .Rows = 2
      .ColWidth(0) = 800
      .ColWidth(1) = 2400
      .ColWidth(2) = 2400
      .ColWidth(3) = 800
      .ColAlignment(0) = 0
      .ColAlignment(1) = 0
      .ColAlignment(2) = 0
      .ColAlignment(3) = 0
      .Row = 0
      .Col = 0
      .Text = "Lots Rq'd"
      .Col = 2
      .Text = "Part Number"
      .Col = 2
      .Text = "Part Description"
      .Col = 3
      .Text = "Level"
      .Row = 1
      .Col = 0
      Set Grid1.CellPicture = Chkno.Picture
   End With
   bRefreshed = 0
   bOnLoad = 1
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveOptions
   
End Sub

Private Sub FillCombo()
   sSql = "Qry_FillSortedParts"
   LoadComboBox cmbPrt
   If cmbPrt.ListCount > 0 Then
      cmbPrt = cmbPrt.List(0)
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "FillCombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set LotsLTe02a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtPrt = "ALL"
   lblInActive.ForeColor = ES_RED
   
End Sub


Private Sub Grid1_KeyPress(KeyAscii As Integer)
   On Error Resume Next
   Grid1.Col = 0
   If lblInActive.Visible = False Then cmdUpd.Enabled = True
   If Grid1.CellPicture = Chkyes.Picture Then
      Set Grid1.CellPicture = Chkno.Picture
      bLtReqd(Grid1.Row) = 0
   Else
      Set Grid1.CellPicture = Chkyes.Picture
      bLtReqd(Grid1.Row) = 1
   End If
   bLtChng(Grid1.Row) = 1
   
End Sub


Private Sub FillGrid()
   Dim RdoGrd As ADODB.Recordset
   Dim b As Byte
   Dim iList As Integer
   Dim a As Integer
   Dim C As Integer
   Dim sPartNumber As String
   Dim sProdCde As String
   
   Grid1.Rows = 1
   Grid1.Row = 0
   Erase sLtPart()
   Erase bLtReqd()
   Erase bLtChng()
   
   On Error GoTo DiaErr1
   a = 10
   On Error GoTo 0
   cmdUpd.Enabled = False
   cmdAll.Enabled = False
   cmdNil.Enabled = False
   If cmbPrt <> "ALL" Then sPartNumber = Compress(cmbPrt)
   If cmbCde <> "ALL" Then sProdCde = Compress(cmbCde)
   
   sSql = "SELECT PARTREF,PARTNUM,PADESC,PAPRODCODE,PALEVEL,PALOTTRACK FROM " _
          & "PartTable WHERE (PARTREF >= '" & sPartNumber & "' AND " _
          & "PAPRODCODE LIKE '" & sProdCde & "%' AND PATOOL=0 "
   If typ(1).Value = vbUnchecked Then sSql = sSql & "AND PartTable.PALEVEL<>1 "
   If typ(2).Value = vbUnchecked Then sSql = sSql & "AND PartTable.PALEVEL<>2 "
   If typ(3).Value = vbUnchecked Then sSql = sSql & "AND PartTable.PALEVEL<>3 "
   If typ(4).Value = vbUnchecked Then sSql = sSql & "AND PartTable.PALEVEL<>4 "
   If typ(5).Value = vbUnchecked Then sSql = sSql & "AND PartTable.PALEVEL<>5 "
   If typ(6).Value = vbUnchecked Then sSql = sSql & "AND PartTable.PALEVEL<>6 "
   If typ(7).Value = vbUnchecked Then sSql = sSql & "AND PartTable.PALEVEL<>7 "
   If typ(8).Value = vbUnchecked Then sSql = sSql & "AND PartTable.PALEVEL<>8 "
   sSql = sSql & ")"
   MouseCursor 13
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoGrd, ES_FORWARD)
   If bSqlRows Then
      Height = 6165
      lblInfo = "Filling Grid."
      lblInfo.Visible = True
      lblInfo.Refresh
      prg1.Value = a
      prg1.Visible = True
      
      With RdoGrd
         Do Until .EOF
            Grid1.Enabled = True
            iList = iList + 1
            C = C + 1
            If C = 10 Then
               a = a + 7
               C = 0
               If a > 95 Then a = 95
               prg1.Value = a
            End If
            sLtPart(iList) = "" & Trim(!PartRef)
            bLtReqd(iList) = !PALOTTRACK
            bLtChng(iList) = 0
            bLtOrig(iList) = bLtReqd(iList)
            
            Grid1.Rows = iList + 1
            
            Grid1.Col = 0
            Grid1.Row = iList
            Grid1.Text = ""
            
            Grid1.Col = 1
            Grid1.Text = "" & Trim(!PartNum)
            
            Grid1.Col = 2
            Grid1.Text = "" & Trim(!PADESC)
            
            Grid1.Col = 3
            Grid1.Text = Format$(!PALEVEL)
            
            Grid1.Col = 0
            If !PALOTTRACK = 1 Then
               ' Grid1.Text = "X"
               Set Grid1.CellPicture = Chkyes.Picture
            Else
               ' Grid1.Text = " "
               Set Grid1.CellPicture = Chkno.Picture
            End If
            lblTotal = iList
            lblTotal.Refresh
            If iList > 299 Then Exit Do
            .MoveNext
         Loop
         ClearResultSet RdoGrd
      End With
      prg1.Value = 100
      Grid1.Col = 0
      Grid1.Row = 1
      On Error Resume Next
      cmdAll.Enabled = True
      cmdNil.Enabled = True
      Grid1.SetFocus
   Else
      MouseCursor 0
      lblTotal = 0
      MsgBox "No Open Operations Found.", _
         vbInformation, Caption
   End If
   prg1.Visible = False
   lblInfo.Visible = False
   Height = 5970
   MouseCursor 0
   Set RdoGrd = Nothing
   Exit Sub
   
DiaErr1:
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub



Private Sub Grid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Grid1.Row > 0 Then Grid1.Col = 0
   On Error Resume Next
   If lblInActive.Visible = False Then cmdUpd.Enabled = True
   If Grid1.CellPicture = Chkyes.Picture Then
      Set Grid1.CellPicture = Chkno.Picture
      bLtReqd(Grid1.Row) = 0
   Else
      Set Grid1.CellPicture = Chkyes.Picture
      bLtReqd(Grid1.Row) = 1
   End If
   bLtChng(Grid1.Row) = 1
   
End Sub

Private Sub txtPrt_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF4 Then
      ViewParts.lblControl = "TXTPRT"
      ViewParts.txtPrt = txtPrt
      optVew.Value = vbChecked
      ViewParts.Show
   End If
   
End Sub

Private Sub txtPrt_LostFocus()
   txtPrt = CheckLen(txtPrt, 30)
   If txtPrt = "" Then txtPrt = "ALL"
   
End Sub

Private Sub cmbPrt_LostFocus()
   cmbPrt = CheckLen(cmbPrt, 30)
   If cmbPrt = "" Then cmbPrt = "ALL"
   
End Sub

Private Sub UpdateList()
   Dim b As Byte
   Dim a As Integer
   Dim iList As Integer
   
   On Error Resume Next
   'Test and update list
   For iList = 1 To Grid1.Rows - 1
      If bLtChng(iList) = 1 Then
         b = 0
         If bLtReqd(iList) <> bLtOrig(iList) Then
            'Parts
            clsADOCon.ADOErrNum = 0
            clsADOCon.BeginTrans
            sSql = "UPDATE PartTable SET PALOTTRACK=" _
                   & bLtReqd(iList) & " WHERE PARTREF='" _
                   & sLtPart(iList) & "' "
            clsADOCon.ExecuteSQL sSql
            If clsADOCon.ADOErrNum = 0 Then
               a = a + 1
               clsADOCon.CommitTrans
            Else
               b = 1
               clsADOCon.RollbackTrans
               clsADOCon.ADOErrNum = 0
            End If
         End If
      End If
   Next
   MsgBox str(a) & " Part Number(s) Were Updated To Reflect " & vbCr _
              & "Changes To The Of Lot Tracking Status.", vbInformation, Caption
   For iList = 1 To Grid1.Rows - 1
      Grid1.Row = iList
      Set Grid1.CellPicture = Chkno.Picture
   Next
   Grid1.Rows = 2
   Grid1.Row = 1
   Grid1.Col = 0
   Grid1.Text = " "
   Grid1.Col = 1
   Grid1.Text = " "
   Grid1.Col = 2
   Grid1.Text = " "
   Grid1.Col = 3
   Grid1.Text = " "
   
   
   Erase sLtPart()
   Erase bLtReqd()
   Erase bLtChng()
   Grid1.Enabled = False
   cmdUpd.Enabled = False
   cmdAll.Enabled = False
   cmdNil.Enabled = False
   On Error Resume Next
   txtPrt.SetFocus
   Exit Sub
   
DiaErr1:
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub SaveOptions()
   Dim iList As Integer
   Dim sOptions As String
   'Save by Menu Option
   For iList = 1 To 8
      sOptions = sOptions & Trim(str(typ(iList).Value))
   Next
   SaveSetting "Esi2000", "EsiQual", "OdSet", Trim(sOptions)
   
End Sub

Private Sub GetOptions()
   Dim iList As Integer
   Dim sOptions As String
   'Get By Menu Option
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiQual", "OdSet", sOptions)
   If Len(sOptions) > 0 Then
      For iList = 1 To 8
         typ(iList) = Mid$(sOptions, iList, 1)
      Next
   End If
   
End Sub

