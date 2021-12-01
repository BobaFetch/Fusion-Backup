VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaGLf02a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Copy a Journal"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6225
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   6225
   ShowInTaskbar   =   0   'False
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   4440
      Top             =   1440
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2385
      FormDesignWidth =   6225
   End
   Begin VB.CheckBox chkAmts 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2160
      TabIndex        =   2
      Top             =   1800
      Width           =   615
   End
   Begin VB.ComboBox cmbJrn 
      Height          =   315
      Left            =   2160
      TabIndex        =   0
      ToolTipText     =   "Includes Only GL Journals Marked As Templates  "
      Top             =   960
      Width           =   1815
   End
   Begin VB.TextBox txtnew 
      Height          =   285
      Left            =   2160
      TabIndex        =   1
      Top             =   1440
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
      Top             =   600
      Width           =   855
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   8
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
      PictureUp       =   "diaGLf02a.frx":0000
      PictureDn       =   "diaGLf02a.frx":0146
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Copy Debits and Credits"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   9
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "All Transactions, Account Numbers, And Comments Will Be Copied.  You May Choose To Copy Debit And Credit Amounts."
      Height          =   495
      Index           =   1
      Left            =   480
      TabIndex        =   7
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "New Journal Name"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Original Journal Name"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   1815
   End
End
Attribute VB_Name = "diaGLf02a"
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
' diaGLf02a - Copy a Journal
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

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd





'**********************************************************************************

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdCpy_Click()
    If Len(Trim(cmbJrn)) = 0 Then
        MsgBox "Please Specify a Journal ID to Copy", vbOKOnly
        cmbJrn.SetFocus
        Exit Sub
    End If
    
    If FindJournal(cmbJrn) = 0 Then
        MsgBox "Your Original Journal ID is Invalid. Please Re-Enter.", vbOKOnly
        cmbJrn.SetFocus
        Exit Sub
    End If
    If Len(Trim(txtnew)) = 0 Then
        MsgBox "Please Specify a New Journal ID", vbOKOnly
        txtnew.SetFocus
        Exit Sub
    End If
    
    If FindJournal(txtnew) = 1 Then
        MsgBox "Your New Journal ID Already Exists. Please Re-Enter."
        txtnew.SetFocus
        Exit Sub
    End If
    
    
   CopyGlJrn
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
      FillCombo
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST
   FormatControls
   bOnLoad = True
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   FormUnload
   Set diaGLf02a = Nothing
End Sub

Public Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Private Sub FillCombo()
   Dim rdoJrn As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT DISTINCT GJNAME FROM GjhdTable WHERE GJTEMPLATE = 1"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoJrn)
   If bSqlRows Then
      With rdoJrn
         While Not .EOF
            AddComboStr cmbJrn.hWnd, "" & Trim(!GJNAME)
            .MoveNext
         Wend
      End With
   End If
   If cmbJrn.ListCount > 0 Then
      cmbJrn.ListIndex = 0
   End If
   Set rdoJrn = Nothing
   Exit Sub
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub txtnew_LostFocus()
   txtnew = CheckLen(txtnew, 12)
   txtnew = CheckComments(txtnew)
End Sub

Private Sub CopyGlJrn()
   Dim rdoJrn1 As ADODB.Recordset
   Dim rdoJrn2 As ADODB.Recordset
   Dim sSource As String
   Dim sDest As String
   Dim sMsg As String
   
   On Error GoTo DiaErr1
   MouseCursor 13
   sSource = "" & Trim(cmbJrn)
   sDest = "" & Trim(txtnew)
   
   ' Copy Header
   Err.Clear
   On Error Resume Next
   clsADOCon.BeginTrans
   clsADOCon.ADOErrNum = 0
   sSql = "SELECT * FROM GjhdTable WHERE GJNAME = '" & sSource & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoJrn1)
   
   With rdoJrn1
      sSql = "INSERT INTO GjhdTable (GJNAME,GJDESC,GJPOST,GJOPEN,GJPOSTED) " _
             & "VALUES (" _
             & "'" & sDest & "'," _
             & "'" & !GJDESC & "'," _
             & "'" & !GJPOST & "'," _
             & "'" & !GJOPEN & "'," _
             & "0)"
      clsADOCon.ExecuteSQL sSql
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
         
         clsADOCon.ExecuteSQL sSql
         .MoveNext
      Wend
   End With
   Set rdoJrn2 = Nothing
   
   If clsADOCon.ADOErrNum = 0 Then
      clsADOCon.CommitTrans
      sMsg = "Successfully Copied " & sSource & " To " & sDest & " ."
      MsgBox sMsg, vbInformation, Caption
      txtnew = ""
      cmbJrn.Clear
      FillCombo
      cmbJrn.SetFocus
   Else
      clsADOCon.RollbackTrans
      clsADOCon.ADOErrNum = 0
      sMsg = "Could Not Copy " & sSource & " To " & sDest & " ."
      MsgBox sMsg, vbInformation, Caption
      txtnew.SetFocus
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




