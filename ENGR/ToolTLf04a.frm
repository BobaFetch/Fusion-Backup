VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form ToolTLf04a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Import Custom Tools From Excel"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8460
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   8460
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdImport 
      Caption         =   "Import Custom Tool Data"
      Height          =   360
      Left            =   3120
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1200
      Width           =   2145
   End
   Begin MSComDlg.CommonDialog fileDlg 
      Left            =   3960
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdOpenDia 
      Caption         =   "..."
      Height          =   255
      Left            =   6960
      TabIndex        =   3
      ToolTipText     =   "Browse XML file or Text file"
      Top             =   720
      Width           =   255
   End
   Begin VB.TextBox txtXLFilePath 
      Height          =   285
      Left            =   2160
      TabIndex        =   2
      Tag             =   "3"
      ToolTipText     =   "Select XML file to import"
      Top             =   720
      Width           =   4695
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "ToolTLf04a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Close"
      Height          =   435
      Left            =   6360
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
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
      FormDesignHeight=   1935
      FormDesignWidth =   8460
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Excel (XLSX) File"
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   4
      Top             =   720
      Width           =   1875
   End
End
Attribute VB_Name = "ToolTLf04a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007)) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
'9/1/04 New
'4/26/06 Corrected GetThisTool query
Option Explicit
Dim bCancel As Byte
Dim bGoodTool As Byte
Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   bCancel = 1
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 3452
      MouseCursor 0
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
      MsgBox "Please select an Excel XLSX file.", _
            vbInformation, Caption
      Exit Sub
   End If

   MouseCursor 13
   ImportTools (strFilePath)
   
   MouseCursor 0
   
   Exit Sub
DiaErr1:
   MouseCursor 0
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me

End Sub

Private Sub cmdOpenDia_Click()
   fileDlg.filter = "Excel Files (*.xlsx) | *.xlsx|"
   
   fileDlg.ShowOpen
   If fileDlg.FileName = "" Then
       txtXLFilePath.Text = ""
   Else
       txtXLFilePath.Text = fileDlg.FileName
   End If
End Sub

Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   bOnLoad = 0
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
   Set ToolTLf03a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Function ImportTools(strFullPath As String) As Boolean
   'returns True if successful

   MsgBox "Excel Dims"
   Dim xlApp As Excel.Application
   Dim wb As Workbook
   Dim ws As Worksheet

   On Error GoTo DiaErr1
   ImportTools = False
   If (strFullPath = "") Then Exit Function
   MsgBox "new Excel.Application"
   Set xlApp = New Excel.Application
   MsgBox "open " & strFullPath
   Set wb = xlApp.Workbooks.Open(strFullPath)
   MsgBox "select worksheet"
   Set ws = wb.Worksheets(1) 'Specify your worksheet name
   
   'get valid categories
   MsgBox "select categories"
   Dim rs As ADODB.Recordset
   Dim categories() As String
   Dim catCount As Integer
   sSql = "select ToolCategory from ToolNewCategories"
   bSqlRows = clsADOCon.GetDataSet(sSql, rs, ES_FORWARD)
   If bSqlRows Then
      With rs
         Do Until .EOF
            catCount = catCount + 1
            ReDim Preserve categories(catCount)
            categories(catCount) = !ToolCategory
            .MoveNext
         Loop
      End With
   End If
   ClearResultSet rs
    
   'get column names
   Dim iIndex As Integer
   iIndex = 4
   Dim colCount As Integer
   Dim col As String
   Dim cols() As String
   Dim types() As String
   Dim lengths() As Integer
   Dim colList As String   'for SQL
   For colCount = 1 To ws.Columns.Count + 1
      MsgBox "select col " & colCount
      col = ws.Cells(iIndex, colCount)
      MsgBox "col = " & col
      If col = "" Then
         colCount = colCount - 1
         Exit For
      End If
      
'         If col = "TOOL_GOVPRIMECONTRACT" Then  'one before TOOL_CATEGORY
'            iIndex = iIndex
'         End If

      ReDim Preserve cols(colCount)
      ReDim Preserve types(colCount)
      ReDim Preserve lengths(colCount)
      cols(colCount) = col
      If colCount = 1 Then
         colList = "(" & col
      Else
         colList = colList & "," & col
      End If
      
      'validate column name and get type
      sSql = "select DATA_TYPE, CHARACTER_MAXIMUM_LENGTH from INFORMATION_SCHEMA.COLUMNS" & vbCrLf _
         & "where TABLE_NAME = 'TlnhdTableNew' AND COLUMN_NAME = '" & col & "'"
      bSqlRows = clsADOCon.GetDataSet(sSql, rs, ES_FORWARD)
      If bSqlRows Then
         With rs
            types(colCount) = !DATA_TYPE
            If IsNumeric(!CHARACTER_MAXIMUM_LENGTH) Then
               lengths(colCount) = CInt(!CHARACTER_MAXIMUM_LENGTH)
            Else
               lengths(colCount) = 0
            End If
            ClearResultSet rs
         End With
      Else
         MsgBox "Column " & col & " does not exist!"
         Exit Function
      End If
   Next colCount
   
   If colCount <= 0 Then Exit Function
   
   'add TOOLNUMREF column at the end
   colCount = colCount + 1
   ReDim Preserve cols(colCount)
   ReDim Preserve types(colCount)
   ReDim Preserve lengths(colCount)
   cols(colCount) = "TOOL_NUMREF"
   types(colCount) = types(2)    'same as TOOL_NUMREF
   lengths(colCount) = lengths(2)
   colList = colList & ",TOOL_NUMREF)" & vbCrLf
   
   'now test each remaining row.  if the tool does not exist create it.
   'terminate on a blank TOOL_NUM
   Dim existingTools As String
   Dim insertCount As Integer
   Dim insertedTools As String
   Dim insertsFailed As Integer
   Dim failedtools As String
   Dim values As String
   iIndex = iIndex + 1
   
   Dim cellVal As String
   Dim refVal As String
   Dim validRow As Boolean
   Dim C As Integer
   Do While (iIndex <= ws.Rows.Count)
      validRow = True
      refVal = ""
      If ws.Cells(iIndex, 2) = "" Then Exit Do
      For C = 1 To colCount - 1
         If C = 1 Then
            values = "values("
         Else
            values = values & ","
         End If

         cellVal = Trim(ws.Cells(iIndex, C))
         
         If C = 2 Then
            refVal = Compress(cellVal)
            'Debug.Print refVal
         End If
         'Debug.Print "row " & iIndex & " tool " & refVal & " col " & cols(C)
            
         'varchar or char
         If lengths(C) > 0 Then
            'string
            If Len(cellVal) > lengths(C) Then
               cellVal = Left(cellVal, lengths(C))
            End If
            
            'for TOOL_CATEGORY, only allow valid choices
            If cols(C) = "TOOL_CATEGORY" Then
               Dim j As Integer
               For j = 1 To catCount
                  If cellVal = categories(j) Then Exit For
               Next j
               If j > catCount Then
                  MsgBox "category " & cellVal & " for tool " & refVal & " is invalid"
                  insertsFailed = insertsFailed + 1
                  If failedtools <> "" Then failedtools = failedtools & ","
                  failedtools = failedtools & refVal
                  validRow = False
               End If
            End If
            values = values & "'" & Replace(cellVal, "'", "''") & "'"   'escape apostrophes
         'bit
         ElseIf types(C) = "bit" Then
            If cellVal = "0" Or Len(cellVal) = 0 Then
               cellVal = "0"
            ElseIf UCase(cellVal) = "NO" Or UCase(cellVal) = "FALSE" Then
               cellVal = "0"
            Else
               cellVal = "1"
            End If
            values = values & cellVal
         'date
         Else
            If cellVal = "" Then
               values = values & "null"
            ElseIf IsDate(cellVal) Then
               values = values & "'" & cellVal & "'"
            Else
               MsgBox cellVal & " is not a valid date.  Tool " & refVal & " will not be inserted"
               insertsFailed = insertsFailed + 1
               If failedtools <> "" Then failedtools = failedtools & ","
               failedtools = failedtools & refVal
               validRow = False
            End If
         End If
      Next C
      
      'add TOOL_NUMREF
      values = values & ",'" & refVal & "')"
      
      'check if tool number already exists
      If validRow Then
         sSql = "select TOOL_NUM from TlnhdTableNew where TOOL_NUMREF = '" & refVal & "'"
         bSqlRows = clsADOCon.GetDataSet(sSql, rs, ES_FORWARD)
         If bSqlRows Then
            With rs
               If Len(existingTools) > 0 Then existingTools = existingTools & ","
               existingTools = existingTools & refVal
               validRow = False
               ClearResultSet rs
            End With
         End If
      End If
      
      'if row is valid, insert it
      If validRow Then
         sSql = "insert tlnhdTableNew " & colList & values
         'Debug.Print sSql
         If clsADOCon.ExecuteSql(sSql) Then
            insertCount = insertCount + 1
            If insertedTools <> "" Then insertedTools = insertedTools & ","
            insertedTools = insertedTools & refVal
         Else
            insertsFailed = insertsFailed + 1
            If failedtools <> "" Then failedtools = failedtools & ","
            failedtools = failedtools & refVal

         End If
      End If
      iIndex = iIndex + 1
   Loop
   
      
   'done
   ImportTools = True
   wb.Close
   xlApp.Quit
   Set ws = Nothing
   Set wb = Nothing
   Set xlApp = Nothing
   
   Dim msg As String
   msg = insertCount & " tool" & IIf(insertCount > 1, "s", "") & " inserted" & vbCrLf
   If insertCount > 0 Then
      msg = msg & insertedTools & vbCrLf
   End If
   If insertsFailed > 0 Then
      msg = msg & insertsFailed & " row insert" & IIf(insertsFailed > 1, "s", "") & " failed" & vbCrLf
      msg = msg & failedtools & vbCrLf
   End If
   If existingTools <> "" Then
      msg = msg & "The following tools already existed" & vbCrLf _
         & existingTools
   End If
   MsgBox msg
   Exit Function
   
DiaErr1:
   wb.Close
   xlApp.Quit
   Set ws = Nothing
   Set wb = Nothing
   Set xlApp = Nothing
   sProcName = "ImportTools"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Function


