VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form CyclCYf02a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Update Inventory Dates For ABC Classes"
   ClientHeight    =   3645
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
   ScaleHeight     =   3645
   ScaleWidth      =   6825
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   1800
      TabIndex        =   2
      Top             =   1920
      Width           =   3015
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "CyclCYf02a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   29
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdVew 
      DownPicture     =   "CyclCYf02a.frx":07AE
      Height          =   315
      Left            =   6360
      Picture         =   "CyclCYf02a.frx":0C88
      Style           =   1  'Graphical
      TabIndex        =   28
      TabStop         =   0   'False
      ToolTipText     =   "View Selections"
      Top             =   1920
      Width           =   350
   End
   Begin VB.CheckBox optPrv 
      Alignment       =   1  'Right Justify
      Caption         =   "Update Last"
      Height          =   195
      Left            =   4800
      TabIndex        =   8
      ToolTipText     =   "Update The Last Cycle Count Date To ""From"" Date (Parts With No Date)"
      Top             =   2640
      Width           =   1215
   End
   Begin VB.ComboBox txtDue 
      Height          =   315
      Left            =   4680
      TabIndex        =   3
      Tag             =   "4"
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CheckBox optNull 
      Alignment       =   1  'Right Justify
      Caption         =   "No Date"
      Height          =   195
      Left            =   3240
      TabIndex        =   7
      ToolTipText     =   "Fill Empty Due Dates (Never Recorded)"
      Top             =   2640
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CheckBox optOvr 
      Alignment       =   1  'Right Justify
      Caption         =   "Replace"
      Height          =   255
      Left            =   1800
      TabIndex        =   6
      ToolTipText     =   "Overwrite And Reset Existing Due Dates"
      Top             =   2640
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.ComboBox txtBeg 
      Height          =   315
      Left            =   1800
      TabIndex        =   1
      Tag             =   "4"
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmdUpd 
      Caption         =   "&Apply"
      Height          =   315
      Left            =   5880
      TabIndex        =   9
      ToolTipText     =   "Update Part Number Count Due Dates To The ""Next"" Date"
      Top             =   1200
      Width           =   875
   End
   Begin VB.ComboBox cmbCde 
      Height          =   315
      Left            =   1800
      Sorted          =   -1  'True
      TabIndex        =   5
      Tag             =   "3"
      ToolTipText     =   "Select Product Code From List (Blank For All)"
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CheckBox optVew 
      Caption         =   "View"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   4080
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtPrt 
      Height          =   285
      Left            =   1800
      TabIndex        =   4
      ToolTipText     =   "Leading Characters Or Blank For All"
      Top             =   1920
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.CommandButton cmdFnd 
      DownPicture     =   "CyclCYf02a.frx":1162
      Height          =   315
      Left            =   4920
      Picture         =   "CyclCYf02a.frx":14A4
      Style           =   1  'Graphical
      TabIndex        =   18
      TabStop         =   0   'False
      ToolTipText     =   "Find A Part Number"
      Top             =   1920
      Visible         =   0   'False
      Width           =   350
   End
   Begin VB.ComboBox cmbAbc 
      ForeColor       =   &H00800000&
      Height          =   315
      ItemData        =   "CyclCYf02a.frx":17E6
      Left            =   1800
      List            =   "CyclCYf02a.frx":17F6
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Select From List Or Leave Blank"
      Top             =   840
      Width           =   615
   End
   Begin VB.CheckBox optInit 
      Enabled         =   0   'False
      Height          =   255
      Left            =   4560
      TabIndex        =   12
      Top             =   360
      Width           =   375
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5880
      TabIndex        =   10
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
      FormDesignHeight=   3645
      FormDesignWidth =   6825
   End
   Begin MSComctlLib.ProgressBar prg1 
      Height          =   300
      Left            =   1320
      TabIndex        =   30
      Top             =   3120
      Width           =   3852
      _ExtentX        =   6800
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   27
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label lblDue 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   5760
      TabIndex        =   26
      ToolTipText     =   "High Value For This ABC Class Code"
      Top             =   3120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "From"
      Height          =   285
      Index           =   5
      Left            =   240
      TabIndex        =   25
      Top             =   1560
      Width           =   1515
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Days To Next"
      Height          =   285
      Index           =   4
      Left            =   3720
      TabIndex        =   24
      Top             =   1560
      Width           =   1155
   End
   Begin VB.Label lblDays 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   3240
      TabIndex        =   23
      ToolTipText     =   "Inventory Frequency"
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label lblPrg 
      Caption         =   "Progress"
      Height          =   255
      Left            =   240
      TabIndex        =   22
      ToolTipText     =   "Establishing ABC Class Codes And Values Initializes ABC Functions"
      Top             =   3120
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Code(s)"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   21
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number(s)"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   19
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label lblCHigh 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4560
      TabIndex        =   17
      ToolTipText     =   "High Value For This ABC Class Code"
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label lblCLow 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1800
      TabIndex        =   16
      ToolTipText     =   "Low Value For This ABC Class Code"
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Class High Value"
      Height          =   285
      Index           =   1
      Left            =   3000
      TabIndex        =   15
      Top             =   1200
      Width           =   1515
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Class Low Value"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   14
      Top             =   1200
      Width           =   1515
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "ABC Class"
      Height          =   285
      Index           =   22
      Left            =   240
      TabIndex        =   13
      Top             =   840
      Width           =   1515
   End
   Begin VB.Label lblStatus 
      Caption         =   "Status"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      ToolTipText     =   "Establishing ABC Class Codes And Values Initializes ABC Functions"
      Top             =   360
      Width           =   4095
   End
End
Attribute VB_Name = "CyclCYf02a"
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

Private Sub SaveOptions()
   Dim sOptions As String
   sOptions = Trim(optOvr.Value) & Trim(optNull.Value) _
              & Trim(optPrv.Value)
   SaveSetting "Esi2000", "EsiInvc", "cydte", sOptions
   
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiInvc", "cydte", sOptions)
   If Trim(sOptions) <> "" Then
      optOvr.Value = Val(Mid(sOptions, 1, 1))
      optNull.Value = Val(Mid(sOptions, 2, 1))
      optPrv.Value = Val(Mid(sOptions, 3, 1))
   End If
   
End Sub

'11/16/03

Private Sub GetClassCode()
   Dim RdoCde As ADODB.Recordset
   Dim dDue As Date
   
   On Error GoTo DiaErr1
   If Trim(txtBeg) <> "" Then dDue = Format(txtBeg, "mm/dd/yyyy") _
           Else dDue = Format(ES_SYSDATE, "mm/dd/yyyy")
   sSql = "SELECT COABCCODE,COABCFREQUENCY,COABCLOWCOST," _
          & "COABCHIGHCOST FROM CabcTable WHERE COABCCODE='" _
          & cmbAbc & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCde, ES_FORWARD)
   If bSqlRows Then
      With RdoCde
         'if exists, else leave it alone
         If Not IsNull(.Fields(0)) Then
            lblDays = Format(.Fields(1), "##0")
            cLowCost = Format(.Fields(2), "#####0.00")
            If cLowCost < cCLowCost Then cLowCost = cCLowCost
            cHighCost = Format(.Fields(3), "#####0.00")
            lblCLow = Format(cLowCost, "###,##0.00")
            lblCHigh = Format(cHighCost, "###,##0.00")
            dDue = Format(dDue + Val(lblDays), "mm/dd/yyyy")
            lblDue = Format(dDue, "mm/dd/yyyy")
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
            cCHighCost = Format(.Fields(1), "#0.00")
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
      OpenHelpContext "5452"
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub

Private Sub cmdUpd_Click()
   Dim bResponse As Byte
   Dim sMsg As String
   If optOvr.Value = vbUnchecked And optNull.Value = vbUnchecked Then
      MsgBox "You Must Check at Least One Option.", _
         vbInformation, Caption
      Exit Sub
   End If
   sMsg = "You Have Chosen To Update The Next Inspection Date For" & vbCr _
          & "Part Numbers, With A Product Code Matching The Selection" & vbCr _
          & "And Have The ABC Class Code Established To The Next " & vbCr _
          & "Date (" & lblDue & "). Do You Wish To Continue?"
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
   GetOptions
   bOnLoad = 1
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveOptions
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set CyclCYf02a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   cmbCde = "ALL"
   txtPrt = "ALL"
   txtBeg = Format(ES_SYSDATE, "mm/dd/yyyy")
   
End Sub

Private Sub FillCombo()
   cmbAbc.Clear
   sSql = "Qry_FillABCCombo"
   LoadComboBox cmbAbc
   If cmbAbc.ListCount > 0 Then
      cmbAbc = cmbAbc.List(0)
      GetClassCode
   End If
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


Private Sub lblDue_Change()
   txtDue = lblDue
   
End Sub

Private Sub optNull_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optOvr_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optPrv_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub txtBeg_DropDown()
   ShowCalendarEx Me
   
End Sub


Private Sub txtBeg_LostFocus()
   Dim dDue As Date
   On Error Resume Next
   txtBeg = CheckDateEx(txtBeg)
   dDue = Format(txtBeg, "mm/dd/yyyy")
   dDue = Format(dDue + Val(lblDays), "mm/dd/yyyy")
   lblDue = Format(dDue, "mm/dd/yyyy")
   
End Sub


Private Sub txtDue_DropDown()
   ShowCalendarEx Me
   
End Sub


Private Sub txtDue_LostFocus()
   txtDue = CheckDateEx(txtDue)
   lblDue = txtDue
   
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
   'Here are ALL
   sSql = "SELECT PARTREF,PAABC,PAPRODCODE,PANEXTCYCLEDATE " _
          & "FROM PartTable WHERE (PAABC='" & Trim(cmbAbc) & "' " _
          & "AND PARTREF LIKE '" & sPartNumber & "%' AND " _
          & "PAPRODCODE LIKE '" & sPcode & "%') "
   If optOvr.Value = vbUnchecked And optNull.Value = vbChecked Then
      sSql = sSql & "AND PANEXTCYCLEDATE IS NULL"
   Else
      If optOvr.Value = vbChecked And optNull.Value = vbUnchecked Then
         sSql = sSql & "AND PANEXTCYCLEDATE IS NOT NULL"
      End If
   End If
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
   
   'Using Current SQL Statement
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
            sSql = "UPDATE PartTable SET PANEXTCYCLEDATE='" & lblDue _
                   & "',PAREVDATE='" & Format(ES_SYSDATE, "mm/dd/yyyy") & "' " _
                   & "WHERE PARTREF='" & Trim(!PartRef) & "'"
            clsADOCon.ExecuteSQL sSql
            If optPrv.Value = vbChecked Then
               sSql = "UPDATE PartTable SET PACYCLEDATE='" & txtBeg _
                      & "' WHERE PARTREF='" & Trim(!PartRef) & "' AND " _
                      & "PACYCLEDATE IS NULL"
               clsADOCon.ExecuteSQL sSql
            End If
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


Private Sub ViewSelections()
   Dim RdoSel As ADODB.Recordset
   Dim iRows As Integer
   Dim sPartNumber As String
   Dim sPcode As String
   
   On Error GoTo DiaErr1
   If cmbPrt <> "" And cmbPrt <> "ALL" Then sPartNumber = Compress(cmbPrt)
   If cmbCde <> "" And cmbCde <> "ALL" Then sPcode = Compress(cmbCde)
   'Here are ALL
   sSql = "SELECT PARTREF,PARTNUM,PAABC,PAPRODCODE,PASTDCOST," _
          & "PACYCLEDATE,PANEXTCYCLEDATE FROM PartTable " _
          & "WHERE (PAABC='" & Trim(cmbAbc) & "' AND " _
          & "PARTREF LIKE '" & sPartNumber & "%' AND " _
          & "PAPRODCODE LIKE '" & sPcode & "%') "
   If optOvr.Value = vbUnchecked And optNull.Value = vbChecked Then
      sSql = sSql & "AND PANEXTCYCLEDATE IS NULL "
   Else
      If optOvr.Value = vbChecked And optNull.Value = vbUnchecked Then
         sSql = sSql & "AND PANEXTCYCLEDATE IS NOT NULL "
      End If
   End If
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
                          ViewCycle.grd.Text = Format(!PACYCLEDATE, "mm/dd/yyyy")
            ViewCycle.grd.Col = 3
            If Not IsNull(!PANEXTCYCLEDATE) Then _
                          ViewCycle.grd.Text = Format(!PANEXTCYCLEDATE, "mm/dd/yyyy")
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
