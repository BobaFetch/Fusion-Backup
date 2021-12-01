VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form diaJRf03a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Post Journals"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6450
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   6450
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbFyr 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1200
      Sorted          =   -1  'True
      TabIndex        =   8
      Tag             =   "8"
      ToolTipText     =   "Select Fiscal Year"
      Top             =   360
      Width           =   1095
   End
   Begin VB.ComboBox cmbAct 
      Height          =   315
      Left            =   3960
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ComboBox cmbTyp 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1200
      TabIndex        =   0
      Tag             =   "8"
      ToolTipText     =   "Select Journal Type From List"
      Top             =   720
      Width           =   2775
   End
   Begin VB.CommandButton cmdpst 
      Caption         =   "&Post"
      Enabled         =   0   'False
      Height          =   315
      Left            =   5520
      TabIndex        =   3
      ToolTipText     =   "Post Selected Journal"
      Top             =   720
      Width           =   875
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   3495
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Open Journals Of Selected Type"
      Top             =   1200
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   6165
      _Version        =   393216
      Rows            =   1
      Cols            =   5
      FixedCols       =   0
      WordWrap        =   -1  'True
      HighLight       =   2
      GridLines       =   0
      ScrollBars      =   2
      SelectionMode   =   1
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5520
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   875
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   4
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
      PictureUp       =   "diaJRf03a.frx":0000
      PictureDn       =   "diaJRf03a.frx":0146
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
      FormDesignHeight=   4800
      FormDesignWidth =   6450
   End
   Begin VB.Label P 
      BackStyle       =   0  'Transparent
      Caption         =   "Fiscal Year"
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   9
      Top             =   360
      Width           =   1185
   End
   Begin VB.Image imgdInc 
      Height          =   180
      Left            =   720
      Picture         =   "diaJRf03a.frx":028C
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgInc 
      Height          =   180
      Left            =   360
      Picture         =   "diaJRf03a.frx":02E3
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label lblTyp 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   4080
      TabIndex        =   5
      Top             =   720
      Width           =   375
   End
End
Attribute VB_Name = "diaJRf03a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of                     ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***

'See the UpdateTables prodecure for database revisions


'************************************************************************************
' diaJRf03a - Post Journal
'
' Created: (nth)
' Revions:
'   09/09/03 (jcw) Allow for multiple journals to be post at once (blue check boxs)
'   11/26/03 (nth) Added time journal posting and default labor account combo
'   09/29/04 (nth) Changed posting to it post directly to GL.
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

Private Sub PostJrnl()
   Dim bResponse As Byte
   Dim sMsg As String
   Dim sCurJrn As String
   Dim sJeName As String
   Dim sJournal As String
   Dim sFis As String
   Dim sNum As String
   Dim sTyp As String
   Dim sDesc As String
   Dim sPost As String
   Dim sNotPost As String
   Dim rdoRoll As ADODB.Recordset
   Dim lTran As Long
   Dim lRef As Long
   Dim cDebit As Currency
   Dim cCredit As Currency
   Dim iRef As Integer
   Dim i As Integer
   Dim iRows As Integer
   Dim sPst As String
   Dim sOpen As String
   
   Dim d As Currency
   Dim c As Currency
   Dim PostThisJournal As Boolean
   
   On Error GoTo DiaErr1
   If iTotalChk > 0 Then
      sMsg = "Are You Certain That You Wish To" & vbCr _
             & "Post These Journals ?"
      bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
      If bResponse = vbNo Then
         CancelTrans
         Exit Sub
      End If
      
      iRows = Grid1.rows - 1
      For i = 1 To iRows
         Grid1.Row = i
         If Grid1.CellPicture = imgInc Then
            Grid1.Col = 1: sJournal = Grid1.Text
            If JournalInBalance(sJournal) Then
            
            PostThisJournal = True
            
            ' check for invalid account numbers
            Dim rs As ADODB.Recordset
            If lblTyp = "TJ" Then
               sSql = "select case when rtrim(TCACCT) = '' then '<BLANK>' else RTRIM(TCACCT) end as Account, count(*) as [Count]" & vbCrLf _
                  & "FROM TcitTable tc" & vbCrLf _
                  & "left join GlacTable gl on tc.TCACCT = gl.GLACCTREF" & vbCrLf _
                  & "WHERE TCGLJOURNAL = '" & sJournal & "'" & vbCrLf _
                  & "and GLACCTNO is null" & vbCrLf _
                  & "group by case when rtrim(TCACCT) = '' then '<BLANK>' else RTRIM(TCACCT) end" & vbCrLf _
                  & "order by case when rtrim(TCACCT) = '' then '<BLANK>' else RTRIM(TCACCT) end"
            Else
               sSql = "select case when rtrim(DCACCTNO) = '' then '<BLANK>' else RTRIM(DCACCTNO) end as Account, count(*) as [Count]" & vbCrLf _
                  & "from JritTable jr" & vbCrLf _
                  & "left join GlacTable gl on jr.DCACCTNO  = gl.GLACCTREF" & vbCrLf _
                  & "where DCHEAD = '" & sJournal & "'" & vbCrLf _
                  & "and GLACCTNO is null" & vbCrLf _
                  & "group by case when rtrim(DCACCTNO) = '' then '<BLANK>' else RTRIM(DCACCTNO) end" & vbCrLf _
                  & "order by case when rtrim(DCACCTNO) = '' then '<BLANK>' else RTRIM(DCACCTNO) end"
            End If
            bSqlRows = clsADOCon.GetDataSet(sSql, rs)
            If bSqlRows Then
               Dim msg As String
               msg = "The following account numbers in journal " & sJournal & " do not exist: " & vbCrLf
               With rs
                  Do Until .EOF
                     msg = msg & "    " & !Account & " (" & CStr(!count) & ")" & vbCrLf
                     .MoveNext
                  Loop
               End With
               Set rs = Nothing
               msg = msg & vbCrLf & "Do you wish to post this journal?"
               If MsgBox(msg, vbYesNo) <> vbYes Then
                  PostThisJournal = False
               End If
            End If
            
            If PostThisJournal Then
   
            
            
               Grid1.Col = 2: sOpen = Grid1.Text
               Grid1.Col = 3: sPst = Grid1.Text
               Grid1.Col = 4: sDesc = Grid1.Text
               
               ' Roll up the journal and post it to the GL
               clsADOCon.BeginTrans
               clsADOCon.ADOErrNum = 0
               sSql = "INSERT INTO GjhdTable (GJNAME,GJDESC,GJPOST,GJOPEN,GJPOSTED) " _
                      & "VALUES ('" & sJournal & "','" & sDesc & "','" & sPst & "'," _
                      & "'" & sOpen & "',1)"
               clsADOCon.ExecuteSql sSql
               iRef = 1
               
               If lblTyp = "TJ" Then
                  ' Special consideration for the time journal since it does
                  ' not reside in JritTable
                  'sSql = "SELECT TCACCT, SUM(TCRATE * TCHOURS + TCOHRATE * TCHOURS) " _
                  '       & "FROM TcitTable WHERE (TCGLJOURNAL = '" & sJournal & "') " _
                  '       & "GROUP BY TCACCT"
                         
                  ' Only company setting says to include the OH cost for posting.
                  If (bPostOHrate = True) Then
                     sSql = "SELECT TCACCT, SUM(TCRATE * TCHOURS + TCOHRATE * TCHOURS) " _
                            & "FROM TcitTable WHERE (TCGLJOURNAL = '" & sJournal & "') " _
                            & "GROUP BY TCACCT"
                  Else
                     sSql = "SELECT TCACCT, SUM(TCRATE * TCHOURS) " _
                            & "FROM TcitTable WHERE (TCGLJOURNAL = '" & sJournal & "') " _
                            & "GROUP BY TCACCT"
                  End If

                  
                  bSqlRows = clsADOCon.GetDataSet(sSql, rdoRoll, ES_STATIC)
                  If bSqlRows Then
                     With rdoRoll
                        While Not .EOF
                           cDebit = .Fields(1)
                           cCredit = cCredit + cDebit
                           sSql = "INSERT INTO GjitTable (JINAME,JITRAN,JIREF," _
                                  & "JIACCOUNT,JIDEB,JICRD) VALUES " _
                                  & "('" & sJournal & "',1," _
                                  & iRef & ",'" _
                                  & Trim(.Fields(0)) & "'," _
                                  & Format(cDebit, "0.00") & "," _
                                  & "0)"
                           clsADOCon.ExecuteSql sSql
                           iRef = iRef + 1
                           .MoveNext
                        Wend
                        .Cancel
                        sSql = "INSERT INTO GjitTable (JINAME,JITRAN,JIREF," _
                               & "JIACCOUNT,JIDEB,JICRD) VALUES " _
                               & "('" & sJournal & "',1," _
                               & iRef & ",'" _
                               & Trim(cmbAct) & "'," _
                               & "0," _
                               & Format(cCredit, "0.00") & ")"
                        clsADOCon.ExecuteSql sSql
                     End With
                     Set rdoRoll = Nothing
                  End If
               Else
                  sSql = "SELECT DCACCTNO, SUM(DCDEBIT) AS DEBIT, SUM(DCCREDIT) AS CREDIT " _
                         & "FROM JritTable WHERE (DCHEAD = '" & Trim(sJournal) _
                         & "') GROUP BY DCACCTNO"
                  bSqlRows = clsADOCon.GetDataSet(sSql, rdoRoll, ES_STATIC)
                  If bSqlRows Then
                     With rdoRoll
                        While Not .EOF
                           If IsNull(!debit) Then
                              cDebit = 0
                           Else
                              cDebit = !debit
                           End If
                           If IsNull(!credit) Then
                              cCredit = 0
                           Else
                              
                              cCredit = !credit
                           End If
                           If cCredit > 0 And cDebit > 0 Then
                              ' Make two journal entries for account
                              sSql = "INSERT INTO GjitTable (JINAME,JITRAN,JIREF," _
                                     & "JIACCOUNT,JIDEB,JICRD) VALUES " _
                                     & "('" & sJournal & "',1," _
                                     & iRef & ",'" _
                                     & Trim(!DCACCTNO) & "'," _
                                     & Format(cDebit, "0.00") & "," _
                                     & "0)"
                              clsADOCon.ExecuteSql sSql
                              iRef = iRef + 1
                              
                              sSql = "INSERT INTO GjitTable (JINAME,JITRAN,JIREF," _
                                     & "JIACCOUNT,JIDEB,JICRD) VALUES " _
                                     & "('" & sJournal & "',1," _
                                     & iRef & ",'" _
                                     & Trim(!DCACCTNO) & "'," _
                                     & "0," _
                                     & Format(cCredit, "0.00") & ")"
                              clsADOCon.ExecuteSql sSql
                              iRef = iRef + 1
                           Else
                              sSql = "INSERT INTO GjitTable (JINAME,JITRAN,JIREF," _
                                     & "JIACCOUNT,JIDEB,JICRD) VALUES " _
                                     & "('" & sJournal & "',1," _
                                     & iRef & ",'" _
                                     & Trim(!DCACCTNO) & "'," _
                                     & Format(cDebit, "0.00") & "," _
                                     & Format(cCredit, "0.00") & ")"
                              clsADOCon.ExecuteSql sSql
                              iRef = iRef + 1
                           End If
                           .MoveNext
                           c = c + cCredit
                           d = d + cDebit
                        Wend
                        .Cancel
                     End With
                     Set rdoRoll = Nothing
                  End If
               End If
               
               'sSql = "UPDATE JrhdTable SET MJCLOSED='" & Format(Now, "mm/dd/yy") _
               '    & "' WHERE MJGLJRNL='" & sJournal & "'"
               'clsAdoCon.ExecuteSQL sSQL
               
               If clsADOCon.ADOErrNum = 0 Then
                  clsADOCon.CommitTrans
                  sPost = sPost & sJournal & vbCrLf
               Else
                  clsADOCon.RollbackTrans
                  clsADOCon.ADOErrNum = 0
                  sNotPost = sNotPost & "*" & sJournal & vbCrLf
                  MsgBox Err
               End If
            Else
               sMsg = sJournal & " is out of balance or has invalid account numbers." _
                      & vbCrLf & "It has not been posted."
               MsgBox sMsg, vbInformation, Caption
            End If
         End If
         End If
      Next
      If Trim(sPost) <> "" Then
         
         MsgBox "Journals: " & vbCrLf & vbCrLf & sPost & vbCrLf & "Posted" _
            & " Successfully.", vbInformation, Caption
      End If
      
      If Trim(sNotPost) <> "" Then
         MsgBox "Journals: " & vbCrLf & vbCrLf & sNotPost & vbCrLf & "Were" _
            & " Not Posted", vbInformation, Caption
      End If
      FillJournals 'Refresh
   Else
      MsgBox "Please Select A Journal", vbInformation, Caption
   End If
   Exit Sub
DiaErr1:
   sProcName = "PostJrnl"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub cmbFyr_Click()
   FillJournals
End Sub

Private Sub cmbFyr_LostFocus()
   FillJournals
End Sub

Private Sub cmbTyp_Click()
   lblTyp = sJournals(cmbTyp.ListIndex, 1)
   FillJournals
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
   PostJrnl
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillCombo
      bOnLoad = False
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   Dim i As Integer
   FormLoad Me, ES_DONTLIST
   FormatControls
   sJournals(0, 0) = "Sales Journal"
   sJournals(0, 1) = "SJ"
   sJournals(1, 0) = "Purchases Journal"
   sJournals(1, 1) = "PJ"
   sJournals(2, 0) = "Cash Receipts Journal"
   sJournals(2, 1) = "CR"
   sJournals(3, 0) = "Computer Checks"
   sJournals(3, 1) = "CC"
   sJournals(4, 0) = "External Checks"
   sJournals(4, 1) = "XC"
'   sJournals(5, 0) = "Payroll Labor Journal"
'   sJournals(5, 1) = "PL"
'   sJournals(6, 0) = "Payroll Disbursements Journal"
'   sJournals(6, 1) = "PD"
   sJournals(5, 0) = "Time Journal"
   sJournals(5, 1) = "TJ"
   
   bPostOHrate = True
   GetLaborAccount
   
   sCurrForm = Caption
   bOnLoad = True
End Sub

Private Sub SetUpGrid()
   Grid1.Cols = 5
   Grid1.ColWidth(0) = 500
   Grid1.ColWidth(1) = 1200
   Grid1.ColWidth(2) = 1000
   Grid1.ColWidth(3) = 1000
   Grid1.ColWidth(4) = 5000
   
   Grid1.Row = 0
   Grid1.Col = 0
   Grid1.Text = "Inc."
   Grid1.Col = 1
   Grid1.Text = "Journal"
   Grid1.Col = 2
   Grid1.Text = "Start"
   Grid1.Col = 3
   Grid1.Text = "End"
   Grid1.Col = 4
   Grid1.Text = "Description"
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set diaJRf03a = Nothing
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Public Sub FillCombo()
   Dim bResponse As Byte
   Dim sMsg As String
   Dim i As Integer
   On Error GoTo DiaErr1
   For i = 0 To 5
      AddComboStr cmbTyp.hWnd, sJournals(i, 0)
   Next
   cmbTyp = cmbTyp.List(0)
   lblTyp = sJournals(0, 1)
   bGoodYear = CheckFiscalYear()
   If bGoodYear Then
      FillFiscalYears Me
      FillJournals
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
   ReDim bChecks(0)
   i = 1
   
   On Error GoTo DiaErr1
   
   Grid1.Clear
   Grid1.rows = 1
   SetUpGrid
   sSql = "SELECT MJGLJRNL,MJDESCRIPTION,MJSTART,MJEND,GJNAME FROM JrhdTable " _
          & "LEFT OUTER JOIN GjhdTable ON MJGLJRNL = GJNAME WHERE (MJCLOSED " _
          & "Is Not Null) And (GJNAME Is Null) AND MJTYPE = '" & lblTyp _
          & "' and MJFY = " & cmbFyr
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoLst, ES_FORWARD)
   If bSqlRows Then
      cmdpst.enabled = True
      With RdoLst
         While Not .EOF
            sItem = "" & Chr(9) & !MJGLJRNL & Chr(9) _
                    & Format(!MJSTART, "mm/dd/yyyy") & Chr(9) _
                    & Format(!MJEND, "mm/dd/yyyy") & Chr(9) _
                    & Trim(!MJDESCRIPTION)
            Grid1.AddItem sItem
            If i > 0 Then
               Grid1.Row = i
               Grid1.Col = 0
               Set Grid1.CellPicture = imgdInc
               Grid1.CellPictureAlignment = flexAlignCenterCenter
            End If
            ReDim Preserve bChecks(i)
            bChecks(i) = 0
            i = i + 1
            .MoveNext
         Wend
         .Cancel
      End With
   Else
      cmdpst.enabled = False
   End If
   Set RdoLst = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "filljournals"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub Grid1_Click()
   If iRow > 0 Then
      Grid1.Row = iRow
      If Grid1.Col = 0 And Grid1.rows > 0 Then
         If bChecks(iRow) = 0 Then
            bChecks(iRow) = 1
            iTotalChk = iTotalChk + 1
            Set Grid1.CellPicture = imgInc
         Else
            bChecks(iRow) = 0
            iTotalChk = iTotalChk - 1
            Set Grid1.CellPicture = imgdInc
         End If
      End If
   End If
End Sub

Private Sub Grid1_MouseDown(Button As Integer, Shift As Integer, _
                            X As Single, Y As Single)
   iRow = Grid1.Row
End Sub

Private Sub GetLaborAccount()
   Dim rdoAct As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT CODEFLABORACCT, IsNULL(COLABOROHTOGL, 0) FROM ComnTable WHERE COREF=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoAct)
   If bSqlRows Then
      With rdoAct
         cmbAct = "" & Trim(.Fields(0))
         bPostOHrate = IIf(Val(.Fields(1) = 0), True, False)
      End With
   End If
   Set rdoAct = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getlaboracct"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub
