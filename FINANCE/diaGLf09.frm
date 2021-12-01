VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form diaGLf09 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Import a General Journal from Excel"
   ClientHeight    =   3330
   ClientLeft      =   1845
   ClientTop       =   615
   ClientWidth     =   8730
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   8730
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtDesc 
      Height          =   315
      Left            =   4680
      MaxLength       =   30
      TabIndex        =   11
      Top             =   1500
      Width           =   3615
   End
   Begin VB.TextBox txtJournalName 
      Height          =   315
      Left            =   1620
      TabIndex        =   9
      Top             =   1500
      Width           =   1635
   End
   Begin VB.ComboBox txtPostDate 
      Height          =   315
      Left            =   1620
      TabIndex        =   6
      Tag             =   "4"
      Top             =   240
      Width           =   1335
   End
   Begin VB.TextBox txtXLFilePath 
      Height          =   285
      Left            =   1620
      TabIndex        =   4
      Tag             =   "3"
      ToolTipText     =   "Select XML file to import"
      Top             =   900
      Width           =   6135
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "Create Journal"
      Height          =   360
      Left            =   2820
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2760
      Width           =   2145
   End
   Begin VB.CommandButton cmdOpenDia 
      Caption         =   "..."
      Height          =   255
      Left            =   8100
      TabIndex        =   2
      ToolTipText     =   "Browse XML file or Text file"
      Top             =   900
      Width           =   255
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   900
      Picture         =   "diaGLf09.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   2760
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   1200
      Top             =   2700
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3330
      FormDesignWidth =   8730
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Close"
      Height          =   435
      Left            =   6780
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   915
   End
   Begin MSComDlg.CommonDialog fileDlg 
      Left            =   240
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Open XLSX File for Import"
      Filter          =   "*.xlsx"
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Descrption"
      Height          =   285
      Index           =   2
      Left            =   3480
      TabIndex        =   12
      Top             =   1560
      Width           =   945
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Journal Name"
      Height          =   285
      Index           =   0
      Left            =   300
      TabIndex        =   10
      Top             =   1560
      Width           =   1185
   End
   Begin VB.Label lblSummary 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1620
      TabIndex        =   8
      Top             =   2160
      Width           =   6705
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Post Date"
      Height          =   285
      Index           =   7
      Left            =   300
      TabIndex        =   7
      Top             =   285
      Width           =   1185
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Excel File"
      Height          =   285
      Index           =   1
      Left            =   300
      TabIndex        =   5
      Top             =   900
      Width           =   1305
   End
End
Attribute VB_Name = "diaGLf09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'Added ITINVOICE

Option Explicit
Dim bCutOff As Byte
Dim bOnLoad As Byte
Dim bUnload As Boolean
Dim debits As Currency
Dim credits As Currency
Dim csv As String
Dim JournalName As String
Dim sMsg As String
Dim rows As Integer
Dim errorflag As Boolean
Private txtKeyPress As New EsiKeyBd
Dim debcrederror As String

Private Sub cmdCan_Click()
   Unload Me

End Sub


Private Sub cmdHlp_Click()
    If cmdHlp Then
        MouseCursor (13)
        OpenHelpContext (2150)
        MouseCursor (0)
        cmdHlp = False
    End If

End Sub

Private Sub cmdImport_Click()
   Dim strWindows As String
   Dim strAccFileName As String
   Dim strFilePath As String
   
   If debcrederror <> "" Then
      SetErrorFlag True, debcrederror
      MsgBox lblSummary
      Exit Sub
   End If
   
   If Not JournalNameOK Then
      MsgBox lblSummary
      Exit Sub
   End If
   
   On Error GoTo DiaErr1
   strFilePath = txtXLFilePath.Text
   
   If (Trim(strFilePath) = "") Then
      MsgBox "Please select a Excel file.", _
            vbInformation, Caption
      Exit Sub
   End If

   If Not ParametersOK Then
      Exit Sub
   End If
   
   MouseCursor 13
   
   'call stored procedure
   Debug.Print Len(csv)
   sSql = "exec InsertGeneralJournal '" & JournalName & "','" & Trim(Replace(txtDesc, "'", "''")) & "','" & txtPostDate & "','" & Replace(csv, "'", "''") & "','" & Secure.UserInitials & "'"
   Dim rs As ADODB.Recordset
   Dim result As String
   Clipboard.Clear
   Clipboard.SetText sSql
   bSqlRows = clsADOCon.GetDataSet(sSql, rs)
   If bSqlRows Then
      With rs
         While Not .EOF
            result = .Fields(0)
            .MoveNext
         Wend
      End With
   End If
   Set rs = Nothing
         
   MouseCursor 0
   
   If InStr(1, result, "created") > 0 Then
      SetErrorFlag False, result
   Else
      SetErrorFlag True, result
   End If
   MsgBox result
   Exit Sub
DiaErr1:
   MouseCursor 0
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me

End Sub

Private Function ParametersOK() As Boolean
   If debcrederror <> "" Then
      ParametersOK = False
      MsgBox debcrederror
      Exit Function
   End If
   
   If Not JournalNameOK Then
      ParametersOK = False
      MsgBox lblSummary
      Exit Function
   End If
   
   If Not PostDateOK Then
      ParametersOK = False
      MsgBox lblSummary
      Exit Function
   End If
   
   ParametersOK = True
   
End Function

Private Sub cmdOpenDia_Click()
   fileDlg.Filter = "Excel Files (*.xlsx) | *.xlsx|"
   
   fileDlg.ShowOpen
   If fileDlg.filename = "" Then
       txtXLFilePath.Text = ""
   Else
       txtXLFilePath.Text = fileDlg.filename
       ReadExcel txtXLFilePath.Text
   End If
End Sub

Private Sub Form_Activate()
   Dim b As Byte
   MdiSect.lblBotPanel = Caption
   
   If bOnLoad Then
      txtPostDate = Format(Now, "MM/dd/yy")
      bOnLoad = 0
   End If
    
   MouseCursor (0)

End Sub

Public Sub MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
  Dim ctl As Control
  
  For Each ctl In Me.Controls
    If TypeOf ctl Is MSFlexGrid Then
      If IsOver(ctl.hWnd, Xpos, Ypos) Then FlexGridScroll ctl, MouseKeys, Rotation, Xpos, Ypos
    End If
  Next ctl
End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   ' make sure that you release the Hook
   'Call WheelUnHook(Me.hWnd)
   
End Sub
Private Sub Form_Load()
    FormLoad Me, ES_DONTLIST
   
   'Call WheelHook(Me.hWnd)
   bOnLoad = 1

End Sub

Private Sub Form_Resize()
    Refresh

End Sub
Private Sub Form_Unload(Cancel As Integer)
    FormUnload
    Set diaGLf09 = Nothing
End Sub


Private Sub ReadExcel(strFullPath As String)

   Dim xlApp As Excel.Application
   Dim wb As Workbook
   Dim ws As Worksheet
   Dim iRow As Integer
   Dim acct As String
   Dim amtString As String
   Dim amtNum As Currency
   debits = 0
   credits = 0
   csv = ""
   errorflag = False
   lblSummary = ""
   debcrederror = ""
   
   On Error GoTo DiaErr1
   
   If (strFullPath <> "") Then
      Set xlApp = New Excel.Application
      Set wb = xlApp.Workbooks.Open(strFullPath)
      Set ws = wb.Worksheets(1)
      Dim debit As Currency, credit As Currency, comment As String
      Dim debitString As String, creditString As String
      
      iRow = 2
      rows = 0
      Do While (True)
         If iRow = 2 Then
            ' get journal and verify it does not exist
            JournalName = Trim(ws.Cells(iRow, 1))
            txtJournalName = JournalName
            JournalNameOK
         End If
         acct = Trim(ws.Cells(iRow, 2))
         debitString = Trim(ws.Cells(iRow, 3))
         creditString = Trim(ws.Cells(iRow, 4))
         comment = Trim(ws.Cells(iRow, 5))
         If JournalName = "" Or acct = "" Or debitString = "" Or creditString = "" Then Exit Do
         If Not AccountExists(acct) Then
            debcrederror = "Account " & acct & " does not exist"
            Exit Do
         End If
         If Not IsNumeric(debitString) Then
             debcrederror = "debit " & debitString & " is not numeric"
             Exit Do
         End If
         If Not IsNumeric(creditString) Then
            debcrederror = "credit " & creditString & " is not numeric"
             Exit Do
         End If
         debit = CCur(debitString)
         credit = CCur(creditString)
         debits = debits + debit
         credits = credits + credit
         If csv <> "" Then csv = csv & ","
         csv = csv & "('" & acct & "'," & CStr(debit) & "," & CStr(credit) & ",'" & comment & "')"
         rows = rows + 1
         iRow = iRow + 1
      Loop
      
      wb.Close
   
      xlApp.Quit
      Set ws = Nothing
      Set wb = Nothing
      Set xlApp = Nothing
      
      If debcrederror = "" Then
         If Not errorflag And (debits <> credits) Then
            debcrederror = "Debits and credits not equal"
         End If
         
         If Not errorflag And (rows < 2) Then
            debcrederror = "At least two debits and credits must be defined"
            Exit Sub
         End If
         
         If Not errorflag And debits = 0 And credits = 0 Then
            debcrederror = "No nonzero debits and credits"
            Exit Sub
         End If
         
         If Not errorflag Then
            SetErrorFlag False, CStr(iRow - 2) & " rows.  debits = " & FormatNumber(debits) & "   credits = " & FormatNumber(credits)
         End If
      Else
         SetErrorFlag True, debcrederror
      End If
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "ReadExcel"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub


Private Sub txtJournalName_LostFocus()
   'verify journal is not in use.
   JournalName = txtJournalName
   If JournalExists(JournalName) Then
      SetErrorFlag True, "Journal " & JournalName & " already exists"
   Else
      If debcrederror = "" Then
         SetErrorFlag False
      End If
   End If
End Sub

Private Sub txtPostDate_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtPostDate_LostFocus()
   txtPostDate = CheckDate(txtPostDate)
   PostDateOK
'   Dim dt As Date, today As Date, minDate As Date, maxDate As Date
'   today = Now
'   minDate = DateAdd("y", -1, today)
'   maxDate = DateAdd("m", 1, today)
'   dt = CDate(txtPostDate)
'   If dt < minDate Or dt > maxDate Then
'      SetErrorFlag True, "Date must be between " & Format(minDate, "MM/dd/yyyy") & " and " & Format(maxDate, "MM/dd/yyyy")
'   Else
'      SetErrorFlag False
'   End If
End Sub

Private Function PostDateOK() As Boolean
   'allow 12 months back or one month in future
   Dim dt As Date, today As Date, minDate As Date, maxDate As Date
   today = Now
   minDate = DateAdd("yyyy", -1, today)
   maxDate = DateAdd("m", 1, today)
   dt = CDate(txtPostDate)
   If dt < minDate Or dt > maxDate Then
      SetErrorFlag True, "Date must be between " & Format(minDate, "MM/dd/yyyy") & " and " & Format(maxDate, "MM/dd/yyyy")
      PostDateOK = False
   Else
      SetErrorFlag False
      PostDateOK = True
   End If

End Function

Private Function JournalExists(JournalName As String) As Boolean
   sSql = "select GJNAME from GjhdTable where GJNAME = '" & JournalName & "'"
   Dim rs As ADODB.Recordset
   JournalExists = clsADOCon.GetDataSet(sSql, rs)
   Set rs = Nothing
End Function

Private Function AccountExists(acct As String) As Boolean
   sSql = "select GLACCTREF from GlacTable where GLACCTREF = '" & acct & "'"
   Dim rs As ADODB.Recordset
   AccountExists = clsADOCon.GetDataSet(sSql, rs)
   Set rs = Nothing
End Function

Private Sub SetErrorFlag(OnOff As Boolean, Optional msg As String = "")
      
   If OnOff Then
      lblSummary = msg
      lblSummary.ForeColor = vbRed
   Else
      lblSummary = msg
      lblSummary.ForeColor = vbTransparent
   End If
   errorflag = OnOff
End Sub

Private Function JournalNameOK() As Boolean
   'returns True if journal name OK
   JournalNameOK = True
   JournalName = Trim(txtJournalName)
   If Len(JournalName) <= 0 Then
      SetErrorFlag True, "Journal name not specified"
      JournalNameOK = False
   ElseIf Len(JournalName) > 12 Then
      SetErrorFlag True, "Journal name must be <= 12 characters"
      JournalNameOK = False
   ElseIf JournalExists(JournalName) Then
      SetErrorFlag True, "Journal " & JournalName & " already exists"
      JournalNameOK = False
   End If
   If JournalNameOK Then
      SetErrorFlag False, ""
   End If
End Function





