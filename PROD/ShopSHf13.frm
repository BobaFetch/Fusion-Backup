VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form ShopSHf13 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Import MO Priorities from Excel"
   ClientHeight    =   1890
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
   ScaleHeight     =   1890
   ScaleWidth      =   8730
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtXLFilePath 
      Height          =   285
      Left            =   1680
      TabIndex        =   4
      Tag             =   "3"
      ToolTipText     =   "Select XML file to import"
      Top             =   240
      Width           =   4695
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "Import Priorities"
      Height          =   360
      Left            =   2820
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1200
      Width           =   2145
   End
   Begin VB.CommandButton cmdOpenDia 
      Caption         =   "..."
      Height          =   255
      Left            =   6600
      TabIndex        =   2
      ToolTipText     =   "Browse XML file or Text file"
      Top             =   240
      Width           =   255
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   120
      Picture         =   "ShopSHf13.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   1020
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   420
      Top             =   960
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   1890
      FormDesignWidth =   8730
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Close"
      Height          =   435
      Left            =   7680
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   915
   End
   Begin MSComDlg.CommonDialog fileDlg 
      Left            =   960
      Top             =   900
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
      Left            =   1680
      TabIndex        =   6
      Top             =   780
      Width           =   4665
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Excel File"
      Height          =   285
      Index           =   1
      Left            =   360
      TabIndex        =   5
      Top             =   240
      Width           =   1305
   End
End
Attribute VB_Name = "ShopSHf13"
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
Dim bUnLoad As Boolean
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
   Dim usr As String
   usr = Secure.UserInitials
   If Len(usr) = 0 Then
      usr = "???"
   End If

   sSql = "exec UpdateMoPriorities '" & Replace(csv, "'", "''") & "','" & usr & "'"
   Dim rs As ADODB.Recordset
   Dim result As String
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
      lblSummary = result
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
   MDISect.lblBotPanel = Caption
'
'   If bOnLoad Then
'      txtPayrollDate = Format(Now, "MM/dd/yy")
'      bOnLoad = 0
'   End If
    
   MouseCursor (0)

End Sub

'Public Sub MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
'  Dim ctl As Control
'
'  For Each ctl In Me.Controls
'    If TypeOf ctl Is MSFlexGrid Then
'      If IsOver(ctl.hwnd, Xpos, Ypos) Then FlexGridScroll ctl, MouseKeys, Rotation, Xpos, Ypos
'    End If
'  Next ctl
'End Sub
'


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   ' make sure that you release the Hook
   ' Call WheelUnHook(Me.hwnd)
   
End Sub
Private Sub Form_Load()
    FormLoad Me, ES_DONTLIST
   
   'Call WheelHook(Me.hwnd)
   bOnLoad = 1

End Sub

Private Sub Form_Resize()
    Refresh

End Sub
Private Sub Form_Unload(Cancel As Integer)
    FormUnload
    Set ShopSHf13 = Nothing
End Sub


Private Function ReadExcel(strFullPath As String)

   Dim xlApp As Excel.Application
   Dim wb As Workbook
   Dim ws As Worksheet
   Dim iRow As Integer
   Dim part As String
   Dim run As String
   Dim priority As String
   Dim amtString As String
   Dim amtNum As Currency
   debits = 0
   credits = 0
   csv = ""
   
   On Error GoTo DiaErr1
   
   If (strFullPath <> "") Then
      Set xlApp = New Excel.Application
      Set wb = xlApp.Workbooks.Open(strFullPath)
      Set ws = wb.Worksheets(1)
      
      iRow = 2
      Do While (True)

         part = Compress(ws.Cells(iRow, 1))
         run = ws.Cells(iRow, 2)
         priority = ws.Cells(iRow, 3)
         If part = "" Or run = "" Or priority = "" Or Not IsNumeric(run) Or Not IsNumeric(priority) Then Exit Do
         
         If csv <> "" Then csv = csv & ","
         csv = csv & "('" & part & "'," & run & "," & priority & ")"
         iRow = iRow + 1
      Loop
      
      wb.Close
   
      xlApp.Quit
      Set ws = Nothing
      Set wb = Nothing
      Set xlApp = Nothing
      
      lblSummary = CStr(iRow - 2) & " MO's found"
   End If
   Exit Function
   
DiaErr1:
   sProcName = "ReadExcel"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Function


