VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form diaJRf07a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Post Year End Journal"
   ClientHeight    =   9660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13710
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9660
   ScaleWidth      =   13710
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtJrnl 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1680
      TabIndex        =   15
      Tag             =   "2"
      Top             =   2040
      Width           =   1455
   End
   Begin VB.TextBox txtDesc 
      Height          =   285
      Left            =   1680
      TabIndex        =   12
      Tag             =   "2"
      Top             =   2400
      Width           =   3615
   End
   Begin VB.CommandButton cmdSel 
      Caption         =   "&Select"
      Height          =   315
      Left            =   3120
      TabIndex        =   11
      ToolTipText     =   "Select Trail Balance"
      Top             =   1560
      Width           =   1215
   End
   Begin VB.ComboBox txtEnd 
      Height          =   315
      Left            =   1680
      TabIndex        =   8
      Tag             =   "4"
      Top             =   1080
      Width           =   1095
   End
   Begin VB.ComboBox txtBeg 
      Height          =   315
      Left            =   1680
      TabIndex        =   7
      Tag             =   "4"
      Top             =   600
      Width           =   1095
   End
   Begin VB.ComboBox cmbFyr 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1680
      Sorted          =   -1  'True
      TabIndex        =   5
      Tag             =   "8"
      ToolTipText     =   "Select Fiscal Year"
      Top             =   1560
      Width           =   1095
   End
   Begin VB.ComboBox cmbAct 
      Height          =   315
      Left            =   3840
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdpst 
      Caption         =   "&New Journal"
      Enabled         =   0   'False
      Height          =   315
      Left            =   7920
      TabIndex        =   2
      ToolTipText     =   "Post Selected Journal"
      Top             =   720
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   6735
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Open Journals Of Selected Type"
      Top             =   2880
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   11880
      _Version        =   393216
      Rows            =   1
      Cols            =   8
      FixedCols       =   0
      WordWrap        =   -1  'True
      HighLight       =   2
      ScrollBars      =   2
      SelectionMode   =   1
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   7920
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   1215
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   3
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
      PictureUp       =   "diaJRf07a.frx":0000
      PictureDn       =   "diaJRf07a.frx":0146
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6360
      Top             =   960
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   9660
      FormDesignWidth =   13710
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Journal ID"
      Height          =   255
      Index           =   2
      Left            =   600
      TabIndex        =   14
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   13
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Period Ending"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   10
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Period Beginning"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   9
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label P 
      BackStyle       =   0  'Transparent
      Caption         =   "Fiscal Year"
      Height          =   285
      Index           =   1
      Left            =   600
      TabIndex        =   6
      Top             =   1560
      Width           =   1185
   End
   Begin VB.Image imgdInc 
      Height          =   180
      Left            =   720
      Picture         =   "diaJRf07a.frx":028C
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgInc 
      Height          =   180
      Left            =   360
      Picture         =   "diaJRf07a.frx":02E3
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "diaJRf07a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************************************************
' diaJRf07a - Post Journal
'
'************************************************************************************

Option Explicit

Dim bOnLoad As Byte
Dim bGoodYear As Byte
Dim bThisYear As Byte
Dim bChecks() As Byte
Dim iRow As Integer
Dim iTotalChk As Integer
Dim iNextJournal As Integer
Dim sJournals(13, 2) As String
Dim bPostOHrate As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub CreateNewJournal()
   Dim bResponse As Byte
   Dim sMsg As String
   Dim sJournal As String
   Dim sActNum As String
   Dim sNet As String
   Dim sDesc As String
   Dim sEndBal As String
   Dim sStartBal  As String
   Dim lTran As Long
   Dim lRef As Long
   Dim i As Integer
   Dim cDebit As Currency
   Dim cCredit As Currency
   Dim cEndBal As Currency
   Dim iRef As Integer
   Dim iRows As Integer
   Dim sOpen As String
   
   On Error GoTo DiaErr1
   sMsg = "Do you Wish To Create New Journal Entry ?"
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbNo Then
      CancelTrans
      Exit Sub
   End If

   sJournal = txtJrnl.Text
   sDesc = txtDesc.Text
   sOpen = Format(Now, "mm/dd/yy")
   
   ' Roll up the journal and post it to the GL
   clsADOCon.BeginTrans
   clsADOCon.ADOErrNum = 0
   
   If (Not CheckGLJrnl(sJournal)) Then
      sSql = "INSERT INTO GjhdTable (GJNAME,GJDESC,GJOPEN,GJPOSTED) " _
             & "VALUES ('" & sJournal & "','" & sDesc & "','" & sOpen & "',0)"
             
      clsADOCon.ExecuteSQL sSql
      iRef = 1
   Else
      iRef = fMaxRef(sJournal)
   End If
   
      
   iRows = Grid1.Rows - 1
   For i = 1 To iRows
      Grid1.Row = i
      Grid1.Col = 0
      If Grid1.CellPicture = imgInc Then
         Grid1.Col = 1: sActNum = Grid1.Text
         Grid1.Col = 3: sStartBal = Grid1.Text
         Grid1.Col = 6: sNet = Grid1.Text
         Grid1.Col = 7: sEndBal = Grid1.Text
         
         If (sStartBal = "") Or (sEndBal = "") Then
            If (sNet <> "") Then
                cEndBal = CDbl(sNet)
            Else
                cEndBal = 0
            End If
         Else
            cEndBal = CDbl(sEndBal)
         End If
         
         If (cEndBal < 0) Then
            cDebit = cEndBal
            cCredit = 0
         Else
            cDebit = 0
            cCredit = cEndBal
         End If
         
         sSql = "INSERT INTO GjitTable (JINAME,JITRAN,JIREF," _
                & "JIACCOUNT,JIDEB,JICRD) VALUES " _
                & "('" & sJournal & "',1," _
                & iRef & ",'" _
                & Trim(sActNum) & "'," _
                & Format(cDebit, "0.00") & "," _
                & Format(cCredit, "0.00") & ")"
                
         clsADOCon.ExecuteSQL sSql
         iRef = iRef + 1

      End If
   Next
      
   If clsADOCon.ADOErrNum = 0 Then
      clsADOCon.CommitTrans
      MsgBox "Created Year End Journal Entry Successfully.", vbInformation, Caption
   Else
      clsADOCon.RollbackTrans
      clsADOCon.ADOErrNum = 0
      MsgBox "Counldn't Create Year End Journal Entry.", vbInformation, Caption
   End If
      
   Exit Sub
DiaErr1:
   sProcName = "CreateNewJournal"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub cmbFyr_LostFocus()
   
   Dim sJournal As String
   Dim sDesc As String
   Dim strStartDt As String
   Dim strEndDt As String
   
   sJournal = "YE" & Trim(cmbFyr)
   CheckGLJrnl sJournal, sDesc
   
   txtJrnl.Text = sJournal
   txtDesc.Text = sDesc

   FillFYDates cmbFyr, strStartDt, strEndDt
   txtBeg = strStartDt
   txtEnd = strEndDt

End Sub

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      SelectHelpTopic Me, Caption
      cmdHlp = False
      MouseCursor 0
   End If
End Sub

Private Sub cmdPst_Click()
   CreateNewJournal
End Sub

Private Sub cmdSel_Click()
   FillJournals
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillCombo
      SetUpGrid
      
      Dim sJournal As String
      Dim sDesc As String
      
      sJournal = "YE" & Trim(cmbFyr)
      CheckGLJrnl sJournal, sDesc
      
      txtJrnl.Text = sJournal
      txtDesc.Text = sDesc
      
      bOnLoad = False
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   Dim i As Integer
   FormLoad Me, ES_DONTLIST
   FormatControls
   GetOptions
   sCurrForm = Caption
   bOnLoad = True
End Sub

Private Sub SetUpGrid()
   Grid1.Cols = 8
   Grid1.ColAlignment(1) = 0
   Grid1.ColWidth(0) = 500
   Grid1.ColWidth(1) = 1000
   Grid1.ColWidth(2) = 3500
   Grid1.ColWidth(3) = 1400
   Grid1.ColWidth(4) = 1400
   Grid1.ColWidth(5) = 1400
   Grid1.ColWidth(6) = 1400
   Grid1.ColWidth(7) = 1400
   
   Grid1.Row = 0
   Grid1.Col = 0
   Grid1.Text = "Inc."
   Grid1.Col = 1
   Grid1.Text = "Account Num"
   Grid1.Col = 2
   Grid1.Text = "Description"
   Grid1.Col = 3
   Grid1.Text = "Beg. Balance"
   Grid1.Col = 4
   Grid1.Text = "Debit"
   Grid1.Col = 5
   Grid1.Text = "Credit"
   Grid1.Col = 6
   Grid1.Text = "Net"
   Grid1.Col = 7
   Grid1.Text = "End Balance"
   
   
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set diaJRf07a = Nothing
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Public Sub FillCombo()
   Dim bResponse As Byte
   Dim sMsg As String
   Dim strStartDt As String
   Dim strEndDt As String
   
   Dim i As Integer
   On Error GoTo DiaErr1
   bGoodYear = CheckFiscalYear()
   If bGoodYear Then
      FillFiscalYears Me
      FillFYDates cmbFyr, strStartDt, strEndDt
      txtBeg = strStartDt
      txtEnd = strEndDt
   Else
      sMsg = "Fiscal Years Have Not Been Initialized." & vbCr _
             & "Initialize Fiscal Years Now?"
      bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
      If bResponse = vbYes Then
         diaGLe04a.Show
         Unload Me
      Else
         cmdpst.enabled = False
      End If
   End If
   Exit Sub
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Public Function CheckFiscalYear() As Byte
   Dim RdoFyr As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT FYYEAR FROM GlfyTable "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoFyr, ES_FORWARD)
   If bSqlRows Then
      CheckFiscalYear = 1
      RdoFyr.Cancel
   Else
      CheckFiscalYear = 0
   End If
   Set RdoFyr = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "checkfisc"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Function

Public Sub FillJournals()
   Dim sItem As String
   Dim RdoLst As ADODB.Recordset
   Dim i As Integer
   Dim lrows As Long
   Dim lRecCnt As Long
   
   i = 1
   MouseCursor 13
   On Error GoTo DiaErr1
   
   Grid1.Rows = 1
   i = 1
   
   Dim strParameterNames(5) As String
   Dim varParameterValues(5) As Variant
   Dim strStoredProcName As String

   strParameterNames(0) = "StartDate"
   strParameterNames(1) = "EndDate"
   strParameterNames(2) = "StartAcc"
   strParameterNames(3) = "EndAcc"
   strParameterNames(4) = "IncludeInActiveAcc"
   
   varParameterValues(0) = Trim(txtBeg)
   varParameterValues(1) = Trim(txtEnd)
   varParameterValues(2) = "ALL"
   varParameterValues(3) = "ALL"
   varParameterValues(4) = "0"
   
   'sSql = "YearEndGLPost  '" & txtBeg & "', '" & txtEnd & "', 'ALL','ALL', '0'"
   
' revised 10/26/2015 -- didn't compile - MM - Un commented.
   bSqlRows = clsADOCon.ExecuteStoredProcEx("YearEndGLPost", strParameterNames, _
                                 varParameterValues, True, RdoLst)
   ' MM Missing Paramter
'   bSqlRows = clsADOCon.ExecuteStoredProcEx("YearEndGLPost", strParameterNames, _
'      varParameterValues, True, , , , , , , , RdoLst)
      
   If bSqlRows Then
      cmdpst.enabled = True
      With RdoLst
         While Not .EOF
         
            sItem = "" & Chr(9) & Trim(!JIACCOUNT) & Chr(9) _
                    & Trim(!GLDESCR) & Chr(9) _
                    & Trim(!StartingBal) & Chr(9) _
                    & Trim(!GLDEBIT) & Chr(9) _
                    & Trim(!GLCREDIT) & Chr(9) _
                    & Trim(!NetProfit) & Chr(9) _
                    & Trim(!EndingBal)
                    
            Grid1.AddItem sItem
            If i > 0 Then
               Grid1.Row = i
               Grid1.Col = 0
               Set Grid1.CellPicture = imgdInc
               Grid1.CellPictureAlignment = flexAlignCenterCenter
            End If
            i = i + 1
            .MoveNext
         Wend
         .Cancel
      End With
   Else
      cmdpst.enabled = False
   End If
   Set RdoLst = Nothing
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   MouseCursor 0
   sProcName = "filljournals"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub Grid1_Click()
   Grid1.Row = iRow
   If Grid1.Col = 0 And Grid1.Rows > 0 Then
      If (Grid1.CellPicture = imgInc) Then
         Set Grid1.CellPicture = imgdInc
      Else
         Set Grid1.CellPicture = imgInc
      End If
   End If
End Sub

Private Sub Grid1_MouseDown(Button As Integer, Shift As Integer, _
                            X As Single, Y As Single)
   iRow = Grid1.Row
End Sub

Private Sub txtend_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtEnd_LostFocus()
   txtEnd = CheckDate(txtEnd)
End Sub

Private Sub txtBeg_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtBeg_LostFocus()
   txtBeg = CheckDate(txtBeg)
End Sub

Private Sub SaveOptions()
   Dim sOptions As String
   sOptions = Trim(txtBeg.Text) & Trim(txtEnd.Text)
   SaveSetting "Esi2000", "EsiFina", Me.Name, sOptions
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   Dim dToday As Integer
   
   dToday = CInt(Mid(Format(Now, "mm/dd/yy"), 4, 2))
   sOptions = GetSetting("Esi2000", "EsiFina", Me.Name, sOptions)
   
   If Len(Trim(sOptions)) > 0 Then
     
     If dToday < 21 Then
      txtBeg = Mid(sOptions, 1, 8)
      txtEnd = Mid(sOptions, 9, 8)
     Else
      txtBeg = Format(Now, "mm/01/yy")
      txtEnd = GetMonthEnd(txtBeg)
     End If
   
   End If
End Sub

Private Function fMaxRef(strGJName As String) As Integer
   Dim rdoRef As ADODB.Recordset
   
   On Error GoTo DiaErr1
   
   ' Get next reference number
   sSql = "SELECT Max(JIREF) + 1 AS MaxOfJIREF FROM GjitTable " _
          & "WHERE JINAME = '" & Trim(strGJName) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoRef, ES_FORWARD)
   
   With rdoRef
      If IsNull(!MaxOfJIREF) Then
         fMaxRef = 1
      Else
         fMaxRef = !MaxOfJIREF
      End If
   End With
   Set rdoRef = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "sfMaxRef"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Function

Private Function CheckGLJrnl(ByRef strGJName As String, Optional ByRef strGJDesc As String) As Boolean
   Dim rdoJrn As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT GJNAME, GJDESC FROM dbo.GjhdTable WHERE GJNAME = '" & strGJName & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoJrn)
   If bSqlRows Then
      With rdoJrn
         strGJDesc = !GJDESC
         CheckGLJrnl = True 'good
      End With
   Else
      strGJDesc = ""
      CheckGLJrnl = False 'bad
   End If
   Set rdoJrn = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "CheckGLJrnl"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Function


Private Sub FillFYDates(strFyYear, ByRef strStartDt As String, ByRef strEndDt As String)
   Dim rdoYr As ADODB.Recordset
   On Error GoTo modErr1
   Dim curFy As String
   curFy = ""
   sSql = "SELECT FYYEAR, FYSTART, FYEND FROM GlfyTable WHERE FYYEAR = '" & strFyYear & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoYr, ES_FORWARD)
   If bSqlRows Then
      With rdoYr
         strStartDt = Format(.Fields(1), "mm/dd/yy")
         strEndDt = Format(.Fields(2), "mm/dd/yy")
      End With
   End If
   Set rdoYr = Nothing
   Exit Sub
modErr1:
   sProcName = "FillFYDates"
   CurrError.Number = Err
   CurrError.Description = Err.Description
End Sub


