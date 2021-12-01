VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form diaGLf08 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Create GL Journal from Payroll Excel Data"
   ClientHeight    =   2745
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
   ScaleHeight     =   2745
   ScaleWidth      =   8730
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox txtPayrollDate 
      Height          =   315
      Left            =   2460
      TabIndex        =   6
      Tag             =   "4"
      Top             =   180
      Width           =   1335
   End
   Begin VB.TextBox txtXLFilePath 
      Height          =   285
      Left            =   2460
      TabIndex        =   4
      Tag             =   "3"
      ToolTipText     =   "Select XML file to import"
      Top             =   840
      Width           =   4695
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "Create Journal"
      Height          =   360
      Left            =   3420
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1980
      Width           =   2145
   End
   Begin VB.CommandButton cmdOpenDia 
      Caption         =   "..."
      Height          =   255
      Left            =   7320
      TabIndex        =   2
      ToolTipText     =   "Browse XML file or Text file"
      Top             =   840
      Width           =   255
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   60
      Picture         =   "diaGLf08.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   60
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   360
      Top             =   0
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2745
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
      Left            =   480
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Open XLSX File for Import"
      Filter          =   "*.xlsx"
   End
   Begin VB.Label lblSummary 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2460
      TabIndex        =   9
      Top             =   1380
      Width           =   4665
   End
   Begin VB.Label lblJournal 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   4200
      TabIndex        =   8
      Top             =   180
      Width           =   1545
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Payroll Date"
      Height          =   285
      Index           =   7
      Left            =   1140
      TabIndex        =   7
      Top             =   225
      Width           =   1185
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Excel File"
      Height          =   285
      Index           =   1
      Left            =   1140
      TabIndex        =   5
      Top             =   840
      Width           =   1305
   End
End
Attribute VB_Name = "diaGLf08"
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

Dim sMsg As String

Private txtKeyPress As New EsiKeyBd


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
   
   On Error GoTo DiaErr1
   strFilePath = txtXLFilePath.Text
   
   If (Trim(strFilePath) = "") Then
      MsgBox "Please select a Excel file.", _
            vbInformation, Caption
      Exit Sub
   End If

   MouseCursor 13
   
   'call stored procedure
   Debug.Print Len(csv)
   sSql = "exec InsertPayrollJournal '" & Replace(csv, "'", "''") & "','" & Secure.UserInitials & "','" & txtPayrollDate & "'"
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
   
   If result <> "" Then
      MsgBox result
   End If
   Exit Sub
DiaErr1:
   MouseCursor 0
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me

End Sub

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
      txtPayrollDate = Format(Now, "MM/dd/yy")
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
   Call WheelUnHook(Me.hWnd)
   
End Sub
Private Sub Form_Load()
    FormLoad Me, ES_DONTLIST
   
   Call WheelHook(Me.hWnd)
   bOnLoad = 1

End Sub

Private Sub Form_Resize()
    Refresh

End Sub
Private Sub Form_Unload(Cancel As Integer)
    FormUnload
    Set diaGLf08 = Nothing
End Sub


Private Function ReadExcel(strFullPath As String)

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
   
   On Error GoTo DiaErr1
   
   If (strFullPath <> "") Then
      Set xlApp = New Excel.Application
      Set wb = xlApp.Workbooks.Open(strFullPath)
      Set ws = wb.Worksheets("Raw Pyrl Data")
      
      iRow = 2
      Do While (True)

         acct = ws.Cells(iRow, 4)
         amtString = ws.Cells(iRow, 5)
         If acct = "" Or amtString = "" Or Not IsNumeric(amtString) Then Exit Do
         
         amtNum = CCur(amtString)
         If amtNum >= 0 Then
            debits = debits + amtNum
         Else
            credits = credits - amtNum
         End If
         
         If csv <> "" Then csv = csv & ","
         csv = csv & "('" & acct & "'," & amtString & ")"
         iRow = iRow + 1
      Loop
      
      wb.Close
   
      xlApp.Quit
      Set ws = Nothing
      Set wb = Nothing
      Set xlApp = Nothing
      
      lblSummary = CStr(iRow - 2) & " rows.  debits = " & FormatNumber(debits) & "   credits = " & FormatNumber(credits)
   End If
   Exit Function
   
DiaErr1:
   sProcName = "ReadExcel"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Function

Private Sub txtPayrollDate_Change()
   If IsDate(txtPayrollDate) Then
      Dim dt As Date
      dt = CDate(txtPayrollDate)
      lblJournal = "PR-" & year(dt) & "-" & Right("0" & Month(dt), 2) & Right("0" & day(dt), 2)
   Else
      lblJournal = ""
   End If
End Sub

Private Sub txtPayrollDate_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtPayrollDate_LostFocus()
   txtPayrollDate = CheckDate(txtPayrollDate)
End Sub

