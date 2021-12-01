VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form diaGLf07a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Copy a Muliple Journal"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6315
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   6315
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox txtOpDt 
      Height          =   315
      Left            =   1320
      TabIndex        =   9
      Tag             =   "4"
      Top             =   960
      Width           =   1335
   End
   Begin VB.ComboBox txtPst 
      Height          =   315
      Left            =   3840
      TabIndex        =   8
      Tag             =   "4"
      Top             =   960
      Width           =   1335
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6000
      Top             =   4080
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   6135
      FormDesignWidth =   6315
   End
   Begin VB.CheckBox chkAmts 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2160
      TabIndex        =   0
      Top             =   1560
      Width           =   255
   End
   Begin VB.TextBox txtnew 
      Height          =   285
      Left            =   3240
      TabIndex        =   2
      Top             =   1560
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5280
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Save And Exit"
      Top             =   120
      Width           =   875
   End
   Begin VB.CommandButton cmdcpy 
      Caption         =   "&Copy"
      Height          =   315
      Left            =   5280
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   600
      Width           =   855
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   6
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
      PictureUp       =   "diaGLf07a.frx":0000
      PictureDn       =   "diaGLf07a.frx":0146
   End
   Begin MSFlexGridLib.MSFlexGrid gridJrl 
      Height          =   4095
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Double Click on the Due Date to Edit"
      Top             =   1920
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   7223
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      ScrollBars      =   2
      AllowUserResizing=   1
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "GL Open Date"
      Height          =   285
      Index           =   7
      Left            =   120
      TabIndex        =   11
      Top             =   1000
      Width           =   1185
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Posting Date"
      Height          =   285
      Index           =   4
      Left            =   2880
      TabIndex        =   10
      Top             =   1000
      Width           =   945
   End
   Begin VB.Image Chkno 
      Height          =   210
      Left            =   5520
      Picture         =   "diaGLf07a.frx":028C
      Stretch         =   -1  'True
      Top             =   1680
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Chkyes 
      Height          =   210
      Left            =   5520
      Picture         =   "diaGLf07a.frx":0616
      Stretch         =   -1  'True
      Top             =   1320
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Copy Debits and Credits"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "All Transactions, Account Numbers, And Comments Will Be Copied.  You May Choose To Copy Debit And Credit Amounts."
      Height          =   495
      Index           =   1
      Left            =   480
      TabIndex        =   5
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "diaGLf07a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of                     ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***
'See the UpdateTables prodecure for database revisions


'**********************************************************************************
'
' diaGLf07a - Copy a Journal
'
' Notes:
'
' Created: 09/30/01 (nth)
' Revisions:
'   08/25/03 (nth) If journal to copy is posted / unpost per incident # 18051
'   04/14/04 (nth) Added template journal filter
'
'**********************************************************************************

Option Explicit

Dim bOnLoad As Byte
Dim bCancel As Byte
Dim sNewJrl As String
Dim bAddNew As Boolean

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd





'**********************************************************************************

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdCpy_Click()
   
   Dim iList As Integer
   Dim strOldJrl As String
   Dim strNewJrl As String
   
   bAddNew = False
   For iList = 1 To gridJrl.Rows - 1
      gridJrl.Row = iList
      
      gridJrl.Col = 1
      strOldJrl = Trim(gridJrl.Text)
      gridJrl.Col = 2
      strNewJrl = Trim(gridJrl.Text)
      
      gridJrl.Col = 0
      If (gridJrl.CellPicture = Chkyes.Picture) And (strNewJrl <> "") Then
         CreateNewJournal strOldJrl, strNewJrl
      End If
   Next
   
   If (bAddNew = True) Then
        MsgBox "Completed creating New Journal ID."
   End If
End Sub


Private Sub CreateNewJournal(strOldJrl As String, strNewJrl As String)
    If Len(Trim(strOldJrl)) = 0 Then
        MsgBox "Please Specify a Journal ID to Copy", vbOKOnly
        Exit Sub
    End If
    
    If FindJournal(strOldJrl) = 0 Then
        MsgBox "Your Original Journal ID is Invalid. Please Re-Enter.", vbOKOnly
        Exit Sub
    End If
    If Len(Trim(strNewJrl)) = 0 Then
        MsgBox "Please Specify a New Journal ID", vbOKOnly
        Exit Sub
    End If
    
    If FindJournal(strNewJrl) = 1 Then
        MsgBox "Your New Journal ID Already Exists. Please Re-Enter."
        Exit Sub
    End If
   
   CopyGlJrn strOldJrl, strNewJrl
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      SelectHelpTopic Me, "No Subject Help"
      cmdHlp = False
      MouseCursor 0
   End If
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      bOnLoad = False
      FillGrid
   End If
   MouseCursor 0
End Sub

Private Sub SetUpGrid()
   With gridJrl
      .Rows = 2
      .Cols = 3
      .RowHeight(0) = 315

      .ColWidth(0) = 500
      .ColWidth(1) = 2500
      .ColWidth(2) = 2500

      .ColAlignment(0) = 0
      .ColAlignment(1) = 1
      .ColAlignment(2) = 1

      .Row = 0
      .Col = 0
      .Text = "Copy"
      .Col = 1
      .Text = "Old Journal Name"
      .Col = 2
      .Text = "New Journal Name"

   End With

End Sub


Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST
   FormatControls
   SetUpGrid
   bOnLoad = True
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   FormUnload
   Set diaGLf07a = Nothing
End Sub

Public Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Private Sub FillGrid()
   Dim rdoJrn As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT DISTINCT GJNAME FROM GjhdTable WHERE GJTEMPLATE = 1"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoJrn)
   gridJrl.Rows = 1
   If bSqlRows Then
      With rdoJrn
         While Not .EOF
            gridJrl.Rows = gridJrl.Rows + 1
            gridJrl.Row = gridJrl.Rows - 1
            
            gridJrl.Col = 0
            Set gridJrl.CellPicture = Chkno.Picture
            gridJrl.Col = 1
            gridJrl.Text = Trim(Compress(!GJNAME))
            
            .MoveNext
         Wend
      End With
   End If
   Exit Sub
DiaErr1:
   sProcName = "FillGrid"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub gridJrl_EnterCell()  ' Assign cell value to the textbox
   txtnew.Text = gridJrl.Text
End Sub

Private Sub txtnew_GotFocus()
   gridJrl.Text = txtnew.Text
   If gridJrl.Col >= gridJrl.Cols Then gridJrl.Col = 2
   'ChangeCellText
End Sub

Private Sub gridJrl_LeaveCell()
   ' Assign textbox value to Grd
   If (txtnew.Visible = True) Then
      gridJrl.Text = txtnew.Text
      ' MM txtnew.Text = ""
      'MM txtnew.Visible = False
   End If

End Sub

Private Sub txtnew_LostFocus()

   If (txtnew.Visible = True) Then
      gridJrl.Text = txtnew.Text
      ' MM txtnew.Text = ""
      ' MM txtnew.Visible = False
   End If
   
End Sub

Public Sub ChangeCellText() ' Move Textbox to active cell.
   txtnew.Move gridJrl.Left + gridJrl.CellLeft, _
   gridJrl.Top + gridJrl.CellTop, _
   gridJrl.CellWidth, gridJrl.CellHeight
   
   'txtnew.SetFocus
End Sub

Private Sub gridJrl_KeyPress(KeyAscii As Integer)
    
   If KeyAscii = vbKeyReturn Or KeyAscii = vbKeySpace Then
      'gridJrl.Col = 0
      If gridJrl.Row >= 1 Then
         If gridJrl.Row = 0 Then gridJrl.Row = 1
         If (gridJrl.Col = 0) Then
            If gridJrl.CellPicture = Chkyes.Picture Then
               Set gridJrl.CellPicture = Chkno.Picture
               gridJrl.Col = 2
               gridJrl.Text = ""
               txtnew.Visible = False
            Else
               Set gridJrl.CellPicture = Chkyes.Picture
               gridJrl.Col = 2
               gridJrl.Text = txtnew.Text
               txtnew.Visible = True
               ChangeCellText
            End If
         ElseIf (gridJrl.Col = 1) Then
            gridJrl.Col = 1
            txtnew.Visible = False
         ElseIf (gridJrl.Col = 2) Then
            'UsingMouse = True
            gridJrl.Col = 2
            gridJrl.Text = txtnew.Text
            txtnew.Visible = True
            ChangeCellText
            
            gridJrl.Col = 0
            If (gridJrl.CellPicture = Chkno.Picture) Then Set gridJrl.CellPicture = Chkyes.Picture
            ' reset back to second column
            gridJrl.Col = 2
         End If
         
      End If
   End If
   

End Sub

Private Sub gridJrl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   'gridJrl.Col = 0
   If gridJrl.Row >= 1 Then
      If gridJrl.Row = 0 Then gridJrl.Row = 1
      If (gridJrl.Col = 0) Then
         If gridJrl.CellPicture = Chkyes.Picture Then
            Set gridJrl.CellPicture = Chkno.Picture
            gridJrl.Col = 2
            gridJrl.Text = ""
            txtnew.Visible = False
         Else
            Set gridJrl.CellPicture = Chkyes.Picture
            gridJrl.Col = 2
            gridJrl.Text = txtnew.Text
            txtnew.Visible = True
            ChangeCellText
         End If
      ElseIf (gridJrl.Col = 1) Then
         gridJrl.Col = 1
         txtnew.Visible = False
      ElseIf (gridJrl.Col = 2) Then
         'UsingMouse = True
         gridJrl.Col = 2
         gridJrl.Text = txtnew.Text
         txtnew.Visible = True
         ChangeCellText
         
         gridJrl.Col = 0
         If (gridJrl.CellPicture = Chkno.Picture) Then Set gridJrl.CellPicture = Chkyes.Picture
         ' reset back to second column
         gridJrl.Col = 2
      End If
      
   End If
End Sub

'Private Sub txtnew_LostFocus()
'   txtnew = CheckLen(txtnew, 12)
'   txtnew = CheckComments(txtnew)
'End Sub

Private Sub CopyGlJrn(sSource As String, sDest As String)
   Dim rdoJrn1 As ADODB.Recordset
   Dim rdoJrn2 As ADODB.Recordset
   Dim sMsg As String
   
   On Error GoTo DiaErr1
   
   If (txtOpDt = "") Or (txtPst = "") Then
      MsgBox "Please Select GL Open Date And Post Date."
      Exit Sub
   End If
   
   MouseCursor 13
   
   ' Copy Header
   Err.Clear
   On Error Resume Next
   clsADOCon.BeginTrans
   clsADOCon.ADOErrNum = 0
   
   sSql = "SELECT * FROM GjhdTable WHERE GJNAME = '" & sSource & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoJrn1, ES_STATIC)
   
   With rdoJrn1
      sSql = "INSERT INTO GjhdTable (GJNAME,GJDESC,GJPOST,GJOPEN,GJPOSTED) " _
             & "VALUES (" _
             & "'" & sDest & "'," _
             & "'" & !GJDESC & "'," _
             & "'" & Format(txtOpDt, "mm/dd/yy") & "'," _
             & "'" & Format(txtPst, "mm/dd/yy") & "'," _
             & "0)"
      
      Debug.Print sSql
      clsADOCon.ExecuteSQL sSql '
   End With
   Set rdoJrn1 = Nothing
   
   ' Copy items
   sSql = "SELECT * FROM GjitTable WHERE JINAME = '" & sSource & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoJrn2, ES_STATIC)
   
   With rdoJrn2
      While Not .EOF
         sSql = "INSERT INTO GjitTable (JINAME,JIDESC,JITRAN,JIREF," _
                & "JIACCOUNT,JIDEB,JICRD) " _
                & "VALUES (" _
                & "'" & sDest & "'," _
                & "'" & Trim(!JIDESC) & "'," _
                & !JITRAN & "," _
                & !JIREF & "," _
                & "'" & Trim(!JIACCOUNT) & "'"
         If chkAmts Then
            sSql = sSql & "," & !JIDEB & "," & !JICRD & ")"
         Else
            sSql = sSql & ",0,0)"
         End If
         
         Debug.Print sSql
         
         clsADOCon.ExecuteSQL sSql '
         
         .MoveNext
      Wend
   End With
   Set rdoJrn2 = Nothing
   
   If clsADOCon.ADOErrNum = 0 Then
      bAddNew = True
      clsADOCon.CommitTrans
   Else
      clsADOCon.RollbackTrans
      clsADOCon.ADOErrNum = 0
      
      sMsg = "Could Not Copy " & sSource & " To " & sDest & " ."
      MsgBox sMsg, vbInformation, Caption
   End If
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "CopyGlJrn"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub



Private Function FindJournal(sJrnID As String) As Byte
   Dim rdoJrn As ADODB.Recordset
   
   sJrnID = Trim(sJrnID)
   FindJournal = 0
   If Len(sJrnID) = 0 Then Exit Function
       
   sSql = "SELECT GJNAME FROM GjhdTable WHERE GJNAME = '" & sJrnID & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoJrn)
   
   On Error Resume Next
   If bSqlRows Then FindJournal = 1
   Set rdoJrn = Nothing
End Function

Private Sub txtOpDt_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtOpDt_LostFocus()
   txtOpDt = CheckDate(txtOpDt)
End Sub

Private Sub txtPst_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtPst_LostFocus()
   txtPst = CheckDate(txtPst)
   If Format(txtOpDt, "yyyy/mm/dd") > Format(txtPst, "yyyy/mm/dd") Then
      txtOpDt = Format(txtPst, "mm/dd/yy")
   End If
End Sub


