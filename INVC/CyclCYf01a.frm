VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form CyclCYf01a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Update ABC Class Codes By Standard Cost"
   ClientHeight    =   3315
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6825
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   6825
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   1800
      TabIndex        =   1
      Top             =   1920
      Width           =   3015
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "CyclCYf01a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   24
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdVew 
      DownPicture     =   "CyclCYf01a.frx":07AE
      Height          =   315
      Left            =   6360
      Picture         =   "CyclCYf01a.frx":0C88
      Style           =   1  'Graphical
      TabIndex        =   22
      TabStop         =   0   'False
      ToolTipText     =   "View Selections"
      Top             =   1920
      Width           =   350
   End
   Begin VB.CommandButton cmdUpd 
      Caption         =   "&Apply"
      Height          =   315
      Left            =   5880
      TabIndex        =   4
      ToolTipText     =   "Update Part Numbers That Have A Standard Cost In The Class Range"
      Top             =   1200
      Width           =   875
   End
   Begin VB.ComboBox cmbCde 
      Height          =   315
      Left            =   1800
      Sorted          =   -1  'True
      TabIndex        =   3
      Tag             =   "3"
      ToolTipText     =   "Select Product Code From List (Blank For All)"
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CheckBox optVew 
      Caption         =   "View"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   3360
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtPrt 
      Height          =   285
      Left            =   1800
      TabIndex        =   2
      ToolTipText     =   "Leading Characters Or Blank For All"
      Top             =   1920
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.CommandButton cmdFnd 
      DownPicture     =   "CyclCYf01a.frx":1162
      Height          =   315
      Left            =   4920
      Picture         =   "CyclCYf01a.frx":14A4
      Style           =   1  'Graphical
      TabIndex        =   17
      TabStop         =   0   'False
      ToolTipText     =   "Find A Part Number"
      Top             =   1920
      Visible         =   0   'False
      Width           =   350
   End
   Begin VB.ComboBox cmbAbc 
      ForeColor       =   &H00800000&
      Height          =   315
      ItemData        =   "CyclCYf01a.frx":17E6
      Left            =   1800
      List            =   "CyclCYf01a.frx":17F6
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Select From List Or Leave Blank"
      Top             =   1200
      Width           =   615
   End
   Begin VB.CheckBox optInit 
      Enabled         =   0   'False
      Height          =   255
      Left            =   4560
      TabIndex        =   11
      Top             =   360
      Width           =   375
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5880
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6600
      Top             =   3840
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3315
      FormDesignWidth =   6825
   End
   Begin MSComctlLib.ProgressBar prg1 
      Height          =   300
      Left            =   1800
      TabIndex        =   25
      Top             =   2760
      Width           =   3852
      _ExtentX        =   6800
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   135
      Left            =   120
      TabIndex        =   23
      Top             =   120
      Width           =   15
   End
   Begin VB.Label lblPrg 
      Caption         =   "Progress"
      Height          =   255
      Left            =   240
      TabIndex        =   21
      ToolTipText     =   "Establishing ABC Class Codes And Values Initializes ABC Functions"
      Top             =   2760
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Code(s)"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   20
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number(s)"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   18
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label lblCHigh 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4560
      TabIndex        =   16
      ToolTipText     =   "High Value For This ABC Class Code"
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label lblCLow 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1800
      TabIndex        =   15
      ToolTipText     =   "Low Value For This ABC Class Code"
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Class High Value"
      Height          =   285
      Index           =   1
      Left            =   3120
      TabIndex        =   14
      Top             =   1560
      Width           =   1515
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Class Low Value"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   13
      Top             =   1560
      Width           =   1515
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "ABC Class"
      Height          =   285
      Index           =   22
      Left            =   240
      TabIndex        =   12
      Top             =   1200
      Width           =   1515
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Your Low Value Is:"
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   10
      ToolTipText     =   "From The Company Setup"
      Top             =   720
      Width           =   2295
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Your High Value Is:"
      Height          =   255
      Index           =   6
      Left            =   3000
      TabIndex        =   9
      ToolTipText     =   "From The Company Setup"
      Top             =   720
      Width           =   2295
   End
   Begin VB.Label lblLow 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1800
      TabIndex        =   8
      ToolTipText     =   "From The Company Setup"
      Top             =   720
      Width           =   975
   End
   Begin VB.Label lblHigh 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4560
      TabIndex        =   7
      ToolTipText     =   "From The Company Setup"
      Top             =   720
      Width           =   975
   End
   Begin VB.Label lblStatus 
      Caption         =   "Status"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      ToolTipText     =   "Establishing ABC Class Codes And Values Initializes ABC Functions"
      Top             =   360
      Width           =   4095
   End
End
Attribute VB_Name = "CyclCYf01a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Stanwood, Washington, USA  ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
'11/17/03 New
Option Explicit
Dim bOnLoad As Byte
Dim iTotalRows As Integer
Dim cCLowCost As Currency
Dim cCHighCost As Currency
Dim cLowCost As Currency
Dim cHighCost As Currency

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub ViewSelections()
   Dim RdoSel As ADODB.Recordset
   Dim iRows As Integer
   Dim sPartNumber As String
   Dim sPcode As String
   
   On Error GoTo DiaErr1
   If cmbPrt <> "" And cmbPrt <> "ALL" Then sPartNumber = Compress(cmbPrt)
   If cmbCde <> "" And cmbCde <> "ALL" Then sPcode = Compress(cmbCde)
   'Here are ALL
   sSql = "SELECT PARTREF,PARTNUM,PASTDCOST,PAPRODCODE," _
          & "PACYCLEDATE,PANEXTCYCLEDATE FROM PartTable WHERE " _
          & "(PASTDCOST BETWEEN " & cLowCost & " AND " & cHighCost & " " _
          & "AND PARTREF LIKE '" & sPartNumber & "%' AND " _
          & "PAPRODCODE LIKE '" & sPcode & "%')"
   sSql = sSql & "ORDER BY PARTREF"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoSel, ES_FORWARD)
   If bSqlRows Then
      With RdoSel
         Do Until .EOF
            iRows = iRows + 1
            If iRows > 1000 Then Exit Do
            ViewCycle.grd.Rows = iRows
            ViewCycle.grd.Row = iRows - 1
            ViewCycle.grd.Col = 0
            ViewCycle.grd.Text = "" & Trim(!PartNum)
            ViewCycle.grd.Col = 1
            ViewCycle.grd.Text = Format(!PASTDCOST, ES_QuantityDataFormat)
            ViewCycle.grd.Col = 2
            If Not IsNull(!PACYCLEDATE) Then _
                          ViewCycle.grd.Text = Format(!PACYCLEDATE, "mm/dd/yy")
            ViewCycle.grd.Col = 3
            If Not IsNull(!PANEXTCYCLEDATE) Then _
                          ViewCycle.grd.Text = Format(!PANEXTCYCLEDATE, "mm/dd/yy")
            .MoveNext
         Loop
         ClearResultSet RdoSel
      End With
      ViewCycle.lblRows = ViewCycle.grd.Rows
      ViewCycle.Caption = "List Of Selected Parts"
      ViewCycle.Show
   Else
      MsgBox "No Matching Part Numbers Were Found.", _
         vbInformation, Caption
   End If
   Set RdoSel = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "viewselect"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

'11/16/03

Private Sub GetClassCode()
   Dim RdoCde As ADODB.Recordset
   
   On Error GoTo DiaErr1
   sSql = "SELECT COABCCODE,COABCLOWCOST,COABCHIGHCOST FROM CabcTable " _
          & "WHERE COABCCODE='" & cmbAbc & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCde, ES_FORWARD)
   If bSqlRows Then
      With RdoCde
         'if exists, else leave it alone
         If Not IsNull(.Fields(0)) Then
            cLowCost = Format(.Fields(1), "#####0.00")
            If cLowCost < cCLowCost Then cLowCost = cCLowCost
            cHighCost = Format(.Fields(2), "#####0.00")
            lblCLow = Format(cLowCost, "###,##0.00")
            lblCHigh = Format(cHighCost, "###,##0.00")
         End If
         ClearResultSet RdoCde
      End With
   End If
   Set RdoCde = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getclassco"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub GetSetup()
   Dim bByte As Byte
   Dim RdoSet As ADODB.Recordset
   
   On Error GoTo DiaErr1
   sSql = "Qry_GetABCPreference"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoSet, ES_FORWARD)
   If bSqlRows Then
      With RdoSet
         If Not IsNull(.Fields(0)) Then
            bByte = .Fields(0)
         Else
            bByte = 0
         End If
         ClearResultSet RdoSet
      End With
   End If
   If bByte = 0 Then
      lblStatus = "The ABC Class Setup Has Not Been Initialized"
      optInit.Value = vbUnchecked
      Set RdoSet = Nothing
      Exit Sub
   Else
      lblStatus = "The ABC Class Setup Has Been Initialized"
      optInit.Value = vbChecked
   End If
   sSql = "SELECT COABCLOWLIMITCOST,COABCHIGHLIMITCOST FROM " _
          & "ComnTable WHERE COREF=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoSet, ES_FORWARD)
   If bSqlRows Then
      With RdoSet
         If Not IsNull(.Fields(0)) Then
            cCLowCost = Format(.Fields(0), "#0.00")
            lblLow = Format(.Fields(0), "#0.00")
            cCHighCost = Format(.Fields(1), "#0.00")
            lblHigh = Format(.Fields(1), "###,##0.00")
         Else
            lblLow = "0.00"
            lblHigh = lblLow
         End If
         ClearResultSet RdoSet
      End With
   End If
   FillCombo
   Set RdoSet = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getsetup"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub cmbAbc_Click()
   GetClassCode
   
End Sub


Private Sub cmbAbc_LostFocus()
   Dim bByte As Byte
   Dim iRows As Integer
   
   For iRows = 0 To cmbAbc.ListCount - 1
      If cmbAbc = cmbAbc.List(iRows) Then bByte = 1
   Next
   If bByte = 0 Then
      Beep
      cmbAbc = cmbAbc.List(0)
   End If
   GetClassCode
   
End Sub


Private Sub cmbCde_LostFocus()
   cmbCde = Trim(cmbCde)
   If cmbCde = "" Then cmbCde = "ALL"
   
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
      OpenHelpContext "5451"
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub

Private Sub cmdUpd_Click()
   Dim bResponse As Byte
   Dim sMsg As String
   sMsg = "You Have Chosen To Update The ABC Class Of The Selected" & vbCr _
          & "Part Numbers, With A Product Code Matching The Selection" & vbCr _
          & "And Have A Standard Cost Between " & lblCLow & " And " & lblCHigh & vbCr _
          & "To The Selected ABC Class Code.  Continue?"
   bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
   If bResponse = vbYes Then
      iTotalRows = GetMatchingParts()
      If iTotalRows = 0 Then
         MsgBox "No Matching Part Numbers Were Found.", _
            vbInformation, Caption
         lblPrg.Visible = False
         prg1.Visible = False
      Else
         sMsg = iTotalRows & " Matching Items Were Found. Do You" & vbCr _
                & "Wish To Continue Updating ABC Classes?"
         bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
         If bResponse = vbYes Then
            UpdateSelections
         Else
            CancelTrans
            lblPrg.Visible = False
            prg1.Visible = False
         End If
      End If
   Else
      CancelTrans
      lblPrg.Visible = False
      prg1.Visible = False
   End If
   cmdUpd.Enabled = True
   
End Sub

Private Sub cmdVew_Click()
   ViewSelections
   
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillProductCodes
      bOnLoad = 0
      GetSetup
      FillComboPart
      cmbPrt = "ALL"
      If optInit.Value = vbUnchecked Then
         MsgBox "You Must First Initialize ABC Class" & vbCr _
            & "Setup Before Using This Function.", vbInformation, Caption
         Unload Me
      End If
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST
   FormatControls
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set CyclCYf01a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   cmbCde = "ALL"
   txtPrt = "ALL"
   
End Sub

Private Sub FillCombo()
   cmbAbc.Clear
   sSql = "Qry_FillABCCombo"
   LoadComboBox cmbAbc
   If cmbAbc.ListCount > 0 Then cmbAbc = cmbAbc.List(0)
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub FillComboPart()
   sSql = "Qry_FillSortedParts"
   LoadComboBox cmbPrt
   If cmbPrt.ListCount > 0 Then
      cmbPrt = cmbPrt.List(0)
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "FillComboPart"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
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
   If Trim(txtPrt) = "" Then txtPrt = "ALL"
   
End Sub

Private Sub cmbPrt_LostFocus()
   cmbPrt = CheckLen(cmbPrt, 30)
   If Trim(cmbPrt) = "" Then cmbPrt = "ALL"
   
End Sub


Private Function GetMatchingParts() As Integer
   Dim RdoSel As ADODB.Recordset
   Dim aProg As Integer
   Dim sPartNumber As String
   Dim sPcode As String
   
   On Error GoTo DiaErr1
   iTotalRows = 0
   prg1.Visible = True
   lblPrg = "Getting Matches"
   lblPrg.Visible = True
   If cmbPrt <> "" And cmbPrt <> "ALL" Then sPartNumber = Compress(cmbPrt)
   If cmbCde <> "" And cmbCde <> "ALL" Then sPcode = Compress(cmbCde)
   sSql = "SELECT PARTREF,PASTDCOST,PAPRODCODE FROM PartTable WHERE " _
          & "(PASTDCOST BETWEEN " & cLowCost & " AND " & cHighCost & " " _
          & "AND PARTREF LIKE '" & sPartNumber & "%' AND " _
          & "PAPRODCODE LIKE '" & sPcode & "%')"
   aProg = aProg + 5
   prg1.Value = aProg
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoSel, ES_FORWARD)
   If bSqlRows Then
      With RdoSel
         Do Until .EOF
            aProg = aProg + 2
            If aProg > 95 Then aProg = 95
            prg1.Value = aProg
            iTotalRows = iTotalRows + 1
            .MoveNext
         Loop
         ClearResultSet RdoSel
      End With
   End If
   GetMatchingParts = iTotalRows
   prg1.Value = 100
   Set RdoSel = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getmatchingpa"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub UpdateSelections()
   Dim RdoSel As ADODB.Recordset
   Dim aProg As Integer
   Dim cCounter As Currency
   Dim sPartNumber As String
   Dim sPcode As String
   
   On Error GoTo DiaErr1
   clsADOCon.ADOErrNum = 0
   
   cmdUpd.Enabled = False
   prg1.Visible = True
   lblPrg = "Updating ABC Classes"
   lblPrg.Visible = True
   If iTotalRows < 1 Then iTotalRows = 1
   cCounter = 95 / iTotalRows
   cCounter = Int(cCounter + 0.4)
   iTotalRows = 0
   If cmbPrt <> "" And cmbPrt <> "ALL" Then sPartNumber = Compress(cmbPrt)
   If cmbCde <> "" And cmbCde <> "ALL" Then sPcode = Compress(cmbCde)
   
   sSql = "SELECT PARTREF,PASTDCOST,PAPRODCODE FROM PartTable WHERE " _
          & "(PASTDCOST BETWEEN " & cLowCost & " AND " & cHighCost & " " _
          & "AND PARTREF LIKE '" & sPartNumber & "%' AND " _
          & "PAPRODCODE LIKE '" & sPcode & "%')"
   aProg = aProg + cCounter
   prg1.Value = aProg
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoSel, ES_FORWARD)
   If bSqlRows Then
      With RdoSel
         On Error Resume Next
         Do Until .EOF
            aProg = aProg + cCounter
            If aProg > 95 Then aProg = 95
            prg1.Value = aProg
            sSql = "UPDATE PartTable SET PAABC='" & Trim(cmbAbc) _
                   & "',PAREVDATE='" & Format(ES_SYSDATE, "mm/dd/yy") & "' " _
                   & "WHERE PARTREF='" & Trim(!PartRef) & "'"
            clsADOCon.ExecuteSQL sSql
            If clsADOCon.ADOErrNum = 0 Then iTotalRows = iTotalRows + 1
            .MoveNext
         Loop
         ClearResultSet RdoSel
      End With
      prg1.Value = 100
      If clsADOCon.ADOErrNum = 0 Then
         MsgBox "Successfully Updated " & iTotalRows & " Parts.", vbInformation, _
            Caption
      Else
         MsgBox "Could Not Successfully Complete the Update.", vbInformation, _
            Caption
      End If
   End If
   Set RdoSel = Nothing
   lblPrg.Visible = False
   prg1.Visible = False
   Exit Sub
   
DiaErr1:
   sProcName = "updateselecti"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub
