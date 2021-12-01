VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form diaJRf02a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Close Journals"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6465
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   6465
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbFyr 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1320
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "8"
      ToolTipText     =   "Select Fiscal Year"
      Top             =   360
      Width           =   1095
   End
   Begin VB.ComboBox cmbTyp 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1320
      TabIndex        =   1
      Tag             =   "8"
      ToolTipText     =   "Select Journal Type From List"
      Top             =   720
      Width           =   2775
   End
   Begin VB.CommandButton cmdCls 
      Caption         =   "&Close Jrn"
      Enabled         =   0   'False
      Height          =   315
      Left            =   5520
      TabIndex        =   3
      ToolTipText     =   "Close This Journal"
      Top             =   720
      Width           =   875
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5520
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   120
      Width           =   875
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
      PictureUp       =   "diaJRf02a.frx":0000
      PictureDn       =   "diaJRf02a.frx":0146
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
      AllowBigSelection=   0   'False
      HighLight       =   2
      GridLines       =   0
      ScrollBars      =   2
      SelectionMode   =   1
   End
   Begin ResizeLibCtl.ReSize ReSize2 
      Left            =   4920
      Top             =   0
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4785
      FormDesignWidth =   6465
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
      Picture         =   "diaJRf02a.frx":028C
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgInc 
      Height          =   180
      Left            =   360
      Picture         =   "diaJRf02a.frx":02E3
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   255
      Index           =   1
      Left            =   4080
      TabIndex        =   8
      Top             =   3840
      Width           =   15
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label lblTyp 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   4200
      TabIndex        =   6
      Top             =   720
      Width           =   375
   End
End
Attribute VB_Name = "diaJRf02a"
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
' diaJRf02a - Close Journals
'
' Created: (nth)
' Revions:
' 09/09/03 (jcw) Allow for multiple journals to be closed at once (blue check boxs)
' 09/28/04 (nth) Added fiscal year filter.
'
'************************************************************************************

Option Explicit

Dim bOnLoad As Byte
Dim bGoodYear As Byte
Dim bThisYear As Byte
Dim iNextJournal As Integer
Dim bChecks() As Byte
Dim bLoadingJrn As Byte
Dim iRow As Integer

Dim sJournals(13, 2) As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd


Private Sub CloseJrn()
   Dim bResponse As Byte
   Dim sMsg As String
   Dim sJrn As String
   Dim sClose As String
   Dim sNotClose As String
   Dim i As Integer
   On Error GoTo DiaErr1
   
'   sJrn = Grid1.Text

   
   sMsg = "Are You Certain That You Wish To" & vbCr _
          & "Close These Journals ?"
   bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
   
   If bResponse = vbYes Then
      Dim CloseDate As Date
      Dim CloseThisJournal As Boolean
      CloseDate = Format(Now, "mm/dd/yyyy")
      For i = 1 To Grid1.rows - 1
         If bChecks(i) Then
            sJrn = Trim(Grid1.TextMatrix(i, 1))
            CloseThisJournal = True
            
            ' check for invalid account numbers
            Dim rs As ADODB.Recordset
            sSql = "select case when rtrim(DCACCTNO) = '' then '<BLANK>' else RTRIM(DCACCTNO) end as Account, count(*) as [Count]" & vbCrLf _
               & "from JritTable jr" & vbCrLf _
               & "left join GlacTable gl on jr.DCACCTNO  = gl.GLACCTREF" & vbCrLf _
               & "where DCHEAD = '" & sJrn & "'" & vbCrLf _
               & "and GLACCTNO is null" & vbCrLf _
               & "group by case when rtrim(DCACCTNO) = '' then '<BLANK>' else RTRIM(DCACCTNO) end" & vbCrLf _
               & "order by case when rtrim(DCACCTNO) = '' then '<BLANK>' else RTRIM(DCACCTNO) end"
            bSqlRows = clsADOCon.GetDataSet(sSql, rs)
            If bSqlRows Then
               Dim msg As String
               msg = "The following account numbers in journal " & sJrn & " do not exist: " & vbCrLf
               With rs
                  Do Until .EOF
                     msg = msg & "    " & !Account & " (" & CStr(!count) & ")" & vbCrLf
                     .MoveNext
                  Loop
               End With
               Set rs = Nothing
               msg = msg & vbCrLf & "Do you wish to close this journal?"
               If MsgBox(msg, vbYesNo) <> vbYes Then
                  CloseThisJournal = False
               End If
            End If
            
            If CloseThisJournal Then
            
            clsADOCon.BeginTrans
            clsADOCon.ADOErrNum = 0
            
            sSql = "UPDATE JrhdTable SET MJCLOSED = '" & CloseDate & "'" & vbCrLf _
               & "WHERE MJGLJRNL ='" & sJrn & "'"
            clsADOCon.ExecuteSql sSql
            
'            'if inventory journal, place journal info in open InvaTable records in date range
'            If Left(sJrn, 3) = "IJ-" Then
'               sSql = "update InvaTable" & vbCrLf
'               sSql = sSql & "set INGLDATE = '" & CloseDate & "'," & vbCrLf
'               sSql = sSql & "INGLJOURNAL = '" & sJrn & "'," & vbCrLf
'               sSql = sSql & "INGLPOSTED = 1" & vbCrLf
'               sSql = sSql & "from InvaTable ia" & vbCrLf
'               sSql = sSql & "join JrhdTable on INADATE between MJSTART and MJEND" & vbCrLf
'               sSql = sSql & "and MJGLJRNL ='" & sJrn & "'" & vbCrLf
'               sSql = sSql & "and INGLDATE is null" & vbCrLf
'               sSql = sSql & "and MJTYPE = 'IJ'"
'               clsADOCon.ExecuteSql sSql
'            End If
'
            'if inventory journal, place journal info in open InvaTable records in date range
            If Left(sJrn, 3) = "IJ-" Then
               sSql = "update InvaTable" & vbCrLf
               sSql = sSql & "set INGLDATE = '" & CloseDate & "'," & vbCrLf
               sSql = sSql & "INGLJOURNAL = '" & sJrn & "'," & vbCrLf
               sSql = sSql & "INGLPOSTED = 1" & vbCrLf
               sSql = sSql & "from InvaTable ia" & vbCrLf
               'sSql = sSql & "join JrhdTable on INADATE between MJSTART and MJEND" & vbCrLf
               sSql = sSql & "join JrhdTable on INADATE >= MJSTART and INADATE < dateadd(day,1,MJEND)" & vbCrLf
               sSql = sSql & "and MJGLJRNL ='" & sJrn & "'" & vbCrLf
               sSql = sSql & "and INGLDATE is null" & vbCrLf
               sSql = sSql & "and MJTYPE = 'IJ'"
               clsADOCon.ExecuteSql sSql
            End If
                        
'            sSql = "UPDATE JrhdTable SET MJCLOSED='" _
'                   & CloseDate & "' " _
'                   & "WHERE MJGLJRNL ='" & Grid1.Text & "'"
'
'
'
'            clsAdoCon.ExecuteSQL sSQL
            If clsADOCon.ADOErrNum = 0 Then
               clsADOCon.CommitTrans
               sClose = sClose & sJrn & vbCrLf
            Else
               clsADOCon.RollbackTrans
               clsADOCon.ADOErrNum = 0
               sNotClose = sNotClose & sJrn & vbCrLf
            End If
            Else
               sNotClose = sNotClose & sJrn & vbCrLf
            End If
         End If
      Next
      If Trim(sClose) <> "" Then
         MsgBox "Journals: " & vbCrLf & vbCrLf & sClose & vbCrLf & "Successfully" _
            & " Closed.", vbInformation, Caption
      End If
      If Trim(sNotClose) <> "" Then
         MsgBox "Journals: " & vbCrLf & vbCrLf & sNotClose & vbCrLf & "Were Not" _
            & " Closed.", vbInformation, Caption
      End If
      
      FillJournals
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "CloseJrnl"
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

Private Sub cmdCls_Click()
   CloseJrn
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      SelectHelpTopic Me, Caption
      cmdHlp = False
      MouseCursor 0
   End If
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
   sJournals(6, 0) = "Inventory Journal"
   sJournals(6, 1) = "IJ"
   
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
   Set diaJRf02a = Nothing
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
   For i = 0 To 6
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
         diaGLe03a.Show
         Unload Me
      Else
         cmdCls.enabled = False
      End If
   End If
   Exit Sub
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub


Private Function CheckFiscalYear() As Byte
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


Private Sub FillJournals()
   Dim sItem As String
   Dim RdoLst As ADODB.Recordset
   On Error GoTo DiaErr1
   
   Dim i As Integer
   ReDim bChecks(0)
   i = 1
   
   Grid1.Clear
   Grid1.rows = 1
   SetUpGrid
   
   sSql = "SELECT MJGLJRNL,MJSTART,MJEND,MJDESCRIPTION " _
          & "FROM JrhdTable WHERE MJTYPE='" & lblTyp & _
          "' AND (MJCLOSED IS NULL) AND MJFY = " & cmbFyr
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoLst)
   If bSqlRows Then
      cmdCls.enabled = True
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
      cmdCls.enabled = False
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
            Set Grid1.CellPicture = imgInc
         Else
            bChecks(iRow) = 0
            Set Grid1.CellPicture = imgdInc
         End If
      End If
   End If
End Sub

Private Sub Grid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   iRow = Grid1.Row
End Sub
