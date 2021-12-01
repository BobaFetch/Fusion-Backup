VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form diaGlJournalEntry
   BorderStyle = 3 'Fixed Dialog
   Caption = "Journal Entry"
   ClientHeight = 6735
   ClientLeft = 45
   ClientTop = 330
   ClientWidth = 8655
   ClipControls = 0 'False
   ControlBox = 0 'False
   LinkTopic = "Form1"
   MaxButton = 0 'False
   MDIChild = -1 'True
   MinButton = 0 'False
   ScaleHeight = 6735
   ScaleWidth = 8655
   ShowInTaskbar = 0 'False
   Begin VB.TextBox txtTrnCrd
      Alignment = 1 'Right Justify
      BackColor = &H8000000F&
      Height = 285
      Left = 3360
      TabIndex = 26
      Tag = "1"
      Top = 1560
      Width = 1200
   End
   Begin VB.TextBox txtTrnDeb
      Alignment = 1 'Right Justify
      BackColor = &H8000000F&
      Height = 285
      Left = 2160
      TabIndex = 25
      Tag = "1"
      Top = 1560
      Width = 1200
   End
   Begin VB.CommandButton cmdDel
      Caption = "&Delete"
      Enabled = 0 'False
      Height = 315
      Left = 7680
      TabIndex = 24
      Top = 6360
      Width = 855
   End
   Begin ResizeLibCtl.ReSize ReSize1
      Left = 8160
      Top = 1560
      _Version = 196615
      _ExtentX = 741
      _ExtentY = 741
      _StockProps = 0
      Enabled = -1 'True
      FormMinWidth = 0
      FormMinHeight = 0
      FormDesignHeight = 6735
      FormDesignWidth = 8655
   End
   Begin VB.TextBox txtDesc
      Height = 285
      Left = 1200
      TabIndex = 1
      Tag = "2"
      Top = 480
      Width = 4455
   End
   Begin VB.ComboBox cmbPost
      Height = 315
      Left = 1200
      TabIndex = 2
      Tag = "4"
      Top = 840
      Width = 1455
   End
   Begin VB.TextBox txtcmt
      Height = 285
      Left = 4680
      TabIndex = 8
      Tag = "2"
      Top = 6000
      Width = 3855
   End
   Begin VB.TextBox txtCrd
      Height = 285
      Left = 3360
      TabIndex = 7
      Tag = "1"
      Top = 6000
      Width = 1200
   End
   Begin VB.TextBox txtDeb
      Height = 285
      Left = 2160
      TabIndex = 6
      Tag = "1"
      Top = 6000
      Width = 1200
   End
   Begin VB.ComboBox cmbAct
      Height = 315
      Left = 720
      TabIndex = 5
      Tag = "3"
      Top = 6000
      Width = 1335
   End
   Begin VB.TextBox txtRef
      BackColor = &H8000000F&
      Enabled = 0 'False
      Height = 285
      Left = 120
      Locked = -1 'True
      TabIndex = 22
      Top = 6000
      Width = 495
   End
   Begin VB.CommandButton cmdUpdate
      Caption = "&Update"
      Height = 315
      Left = 6720
      TabIndex = 9
      Top = 6360
      Width = 855
   End
   Begin VB.CommandButton cmdPost
      Caption = "&Post"
      Height = 315
      Left = 7680
      TabIndex = 21
      ToolTipText = "Post GL Journal Entry"
      Top = 1080
      Width = 855
   End
   Begin VB.ComboBox cmbTran
      Height = 315
      Left = 1200
      TabIndex = 3
      Tag = "1"
      Top = 1560
      Width = 735
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1
      Height = 3615
      Left = 120
      TabIndex = 4
      Top = 2400
      Width = 8415
      _ExtentX = 14843
      _ExtentY = 6376
      _Version = 393216
      Rows = 1
      Cols = 5
      FixedRows = 0
      FixedCols = 0
      BackColor = 16777215
      ScrollBars = 2
      SelectionMode = 1
      FormatString = ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
      Name = "MS Sans Serif"
      Size = 8.25
      Charset = 0
      Weight = 400
      Underline = 0 'False
      Italic = 0 'False
      Strikethrough = 0 'False
      EndProperty
   End
   Begin VB.ComboBox cmbJrn
      Height = 315
      Left = 1200
      Sorted = -1 'True
      TabIndex = 0
      Tag = "2"
      ToolTipText = "Select From List"
      Top = 120
      Width = 1890
   End
   Begin VB.CommandButton cmdCan
      Cancel = -1 'True
      Caption = "Close"
      Height = 435
      Left = 7680
      TabIndex = 10
      TabStop = 0 'False
      Top = 0
      Width = 875
   End
   Begin Threed.SSRibbon cmdHlp
      Height = 225
      Left = 0
      TabIndex = 11
      ToolTipText = "Subject Help"
      Top = 0
      Width = 255
      _Version = 65536
      _ExtentX = 450
      _ExtentY = 397
      _StockProps = 65
      BackColor = 12632256
      Autosize = 2
      RoundedCorners = 0 'False
      BevelWidth = 0
      Outline = 0 'False
      PictureUp = "diaGljrn.frx":0000
      PictureDn = "diaGljrn.frx":0146
   End
   Begin Threed.SSFrame z2
      Height = 30
      Index = 0
      Left = 120
      TabIndex = 12
      Top = 1440
      Width = 8445
      _Version = 65536
      _ExtentX = 14896
      _ExtentY = 53
      _StockProps = 14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
      Name = "MS Sans Serif"
      Size = 8.25
      Charset = 0
      Weight = 400
      Underline = 0 'False
      Italic = 0 'False
      Strikethrough = 0 'False
      EndProperty
   End
   Begin VB.Label txtdiff
      Alignment = 1 'Right Justify
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      Height = 255
      Left = 6360
      TabIndex = 33
      Top = 1080
      Width = 1215
   End
   Begin VB.Label txtDebTotal
      Alignment = 1 'Right Justify
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      Height = 255
      Left = 6360
      TabIndex = 32
      Top = 240
      Width = 1215
   End
   Begin VB.Label txtCrdTotal
      Alignment = 1 'Right Justify
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      Height = 255
      Left = 6360
      TabIndex = 31
      Top = 600
      Width = 1215
   End
   Begin VB.Label lblActDesc
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      Height = 285
      Left = 720
      TabIndex = 30
      Top = 6360
      Width = 3855
   End
   Begin VB.Line Line1
      Index = 1
      X1 = 5520
      X2 = 7560
      Y1 = 960
      Y2 = 960
   End
   Begin VB.Label z1
      Alignment = 1 'Right Justify
      BackStyle = 0 'Transparent
      Caption = "Difference"
      Height = 255
      Index = 10
      Left = 5160
      TabIndex = 29
      Top = 1080
      Width = 1095
   End
   Begin VB.Label z1
      Alignment = 1 'Right Justify
      BackStyle = 0 'Transparent
      Caption = "Credit"
      Height = 255
      Index = 9
      Left = 5400
      TabIndex = 28
      Top = 600
      Width = 855
   End
   Begin VB.Label z1
      Alignment = 1 'Right Justify
      BackStyle = 0 'Transparent
      Caption = "Debit"
      Height = 255
      Index = 1
      Left = 5520
      TabIndex = 27
      Top = 240
      Width = 735
   End
   Begin VB.Label z1
      Alignment = 1 'Right Justify
      BackStyle = 0 'Transparent
      Caption = "Description"
      Height = 255
      Index = 0
      Left = 240
      TabIndex = 23
      Top = 480
      Width = 855
   End
   Begin VB.Label z1
      Alignment = 1 'Right Justify
      BackStyle = 0 'Transparent
      Caption = "Post Date"
      Height = 255
      Index = 11
      Left = 240
      TabIndex = 20
      Top = 840
      Width = 855
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Ref"
      Height = 255
      Index = 8
      Left = 120
      TabIndex = 19
      Top = 2040
      Width = 615
   End
   Begin VB.Label z1
      Alignment = 1 'Right Justify
      BackStyle = 0 'Transparent
      Caption = "Transaction"
      Height = 255
      Index = 7
      Left = 120
      TabIndex = 18
      Top = 1560
      Width = 975
   End
   Begin VB.Line Line1
      Index = 0
      X1 = 120
      X2 = 8520
      Y1 = 2280
      Y2 = 2280
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Credit Amt"
      Height = 255
      Index = 6
      Left = 3240
      TabIndex = 17
      Top = 2040
      Width = 855
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Comments"
      Height = 255
      Index = 5
      Left = 4200
      TabIndex = 16
      Top = 2040
      Width = 1215
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Debit Amt"
      Height = 255
      Index = 4
      Left = 2160
      TabIndex = 15
      Top = 2040
      Width = 855
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Account"
      Height = 255
      Index = 3
      Left = 720
      TabIndex = 14
      Top = 2040
      Width = 975
   End
   Begin VB.Label z1
      Alignment = 1 'Right Justify
      BackStyle = 0 'Transparent
      Caption = "Journal ID"
      Height = 255
      Index = 2
      Left = 240
      TabIndex = 13
      Top = 120
      Width = 855
   End
End
Attribute VB_Name = "diaGlJournalEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************************************
' diaGlJournalEntry - Add/Revise Gl Journal Entries
'
' Created: (nth)
' Revsions:
'
'
'***********************************************************************

Option Explicit

Dim bOnLoad As Byte ' Prevents calling the same event twice during loads
Dim bCancel As Boolean ' Exit form
Dim bGoodId As Byte
Dim bGoodYear As Byte
Dim iFyear As Integer
Dim iJrnNo As Integer
Dim sJrnl As String ' To tell if the journal has changed

' Are we adding or revising

Dim bTrans As Byte

' 1 = AddTran
' 2 = ReviseTran
' 0 = nothing

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Public Sub PostJrnl()
   Dim sMsg As String
   Dim bResponse As Byte
   Dim sJrn As String
   
   On Error GoTo diaErr1
   sJrn = Trim(cmbJrn)
   sMsg = "Do You Wish To Post General Journal " & Trim(cmbJrn) & " ?"
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   
   If bResponse = vbNo Then
      Exit Sub
   End If
   
   On Error Resume Next
   Err = 0
   RdoCon.BeginTrans
   
   sSql = "UPDATE GjhdTable SET " _
          & "GJPOSTED = 1" _
          & " WHERE GJNAME = '" & sJrn & "'"
   RdoCon.Execute sSql, rdExecDirect
   
   If Err = 0 Then
      RdoCon.CommitTrans
      FillJournals
      MouseCursor 0
      MsgBox sJrn & " Successfully Posted.", _
         vbInformation, Caption
   Else
      RdoCon.RollbackTrans
      MouseCursor 0
      MsgBox "Couldn't Successfuly Post " & sJrn & ".", _
         vbExclamation, Caption
   End If
   Exit Sub
   diaErr1:
   sProcName = "PostJrnl"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Function MaxRef() As Integer
   Dim rdoRef As rdoResultset
   
   On Error GoTo diaErr1
   
   ' Get next reference number
   sSql = "SELECT Max(JIREF) AS MaxOfJIREF FROM GjitTable " _
          & "WHERE JINAME = '" & Trim(cmbJrn) & " ' AND JITRAN = " & CInt(cmbTran)
   bSqlRows = GetDataSet(rdoRef, ES_FORWARD)
   
   With rdoRef
      If IsNull(!MaxOfJIREF) Then
         MaxRef = 0
      Else
         MaxRef = !MaxOfJIREF
      End If
   End With
   Set rdoRef = Nothing
   Exit Function
   
   diaErr1:
   sProcName = "MaxRef"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Function

Public Sub FillAccounts()
   ' Fill account combo
   ' Need to add account descriptions
   Dim rdoAct As rdoResultset
   
   sSql = "Qry_FillLowAccounts"
   bSqlRows = GetDataSet(rdoAct, ES_FORWARD)
   
   If bSqlRows Then
      With rdoAct
         Do Until .EOF
            AddComboStr cmbAct.hWnd, "" & Trim(!GLACCTNO)
            .MoveNext
         Loop
      End With
      cmbAct.ListIndex = 0
   End If
   Set rdoAct = Nothing
End Sub

Public Sub OpenJrn()
   Dim RdoJrn As rdoResultset
   Dim iResponse As Integer
   Dim sMsg As String
   Dim iNo As Integer
   Dim iYear As Integer
   
   ' This function is to be called after the cmbJrn looses focus
   On Error GoTo diaErr1
   sJrnl = cmbJrn
   
   sSql = "SELECT GJNAME,GJPOST,GJDESC FROM GjhdTable WHERE GJNAME = '" _
          & Trim(sJrnl) & "'"
   bSqlRows = GetDataSet(RdoJrn, ES_FORWARD)
   
   ' Journal Found!
   If bSqlRows Then
      With RdoJrn
         cmbPost = "" & Trim(!GJPOST)
         txtDesc = "" & Trim(!GJDESC)
      End With
      FillTrans
      cmbTran.ListIndex = cmbTran.ListCount - 1
      
      ' No Journal exists ask user to make one
   Else
      sMsg = "Journal " & sJrnl & " Does Not Exists.  Do You Wish To Create It?"
      iResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
      
      ' Add GL journal to database
      If iResponse = vbYes Then
         
         On Error Resume Next
         RdoCon.BeginTrans
         sSql = "INSERT INTO GjhdTable (GJNAME,GJDESC,GJOPEN,GJPOST,GJPOSTED) " _
                & " VALUES (" _
                & "'" & sJrnl & "'," _
                & "'" & txtDesc & "'," _
                & "'" & Format(Now, "mm/dd/yy") & "'," _
                & "'" & Format(cmbPost, "mm/dd/yy") & "'," _
                & "0)"
         RdoCon.Execute sSql, rdExecDirect
         
         If Err = 0 Then
            RdoCon.CommitTrans
            ' Reset and add first transaction number
            
            cmbTran.Clear
            cmbTran.AddItem "1"
            cmbTran.ListIndex = 0
            
            txtDebTotal = 0# 'Format(0#, "0.00")
            txtCrdTotal = 0# 'Format(0#, "0.00")
            txtdiff = 0# 'Format(0#, "0.00")
            
            FillJournals
            cmbJrn = sJrnl
            MsgBox sJrnl & " Successfully Opened.", vbInformation, Caption
         Else
            
            RdoCon.RollbackTrans
            MouseCursor 0
            MsgBox "Couldn't Successfuly Open " & sJrnl & ".", _
               vbExclamation, Caption
         End If
      End If
   End If
   Set RdoJrn = Nothing
   CalcTotals
   Exit Sub
   diaErr1:
   Set RdoJrn = Nothing
   MouseCursor 0
   sProcName = "OpenJrn"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Public Sub DelTran()
   Dim sMsg As String
   Dim bResponse As Byte
   
   On Error GoTo diaErr1
   Grid1.Col = 0
   sMsg = "Do You Wish To Delete General Journal Entry Ref " & Grid1 & " ?"
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      MouseCursor 13
      On Error Resume Next
      'On Error GoTo 0
      Err = 0
      RdoCon.BeginTrans
      sSql = "DELETE FROM GjitTable WHERE JINAME = '" & Trim(cmbJrn) _
             & "' AND JITRAN = " & CInt(cmbTran) & " AND JIREF = " & CInt(Grid1)
      RdoCon.Execute sSql, rdExecDirect
      If Err = 0 Then
         RdoCon.CommitTrans
         MsgBox "The Transaction Was Completed Successfully.", _
            vbInformation, Caption
         FillTrans
         FillGrid
         CalcTotals
         MouseCursor 0
      Else
         MouseCursor 0
         RdoCon.RollbackTrans
         MsgBox "Couldn't Complete The Transaction.", _
            vbExclamation, Caption
      End If
   End If
   If MaxRef = 0 Then
      cmbTran.ListIndex = 0
   End If
   Grid1.SetFocus
   Exit Sub
   diaErr1:
   MouseCursor 0
   sProcName = "DelTran"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Public Sub FillTrans()
   Dim rdoTran As rdoResultset
   Dim iOldTran As Integer
   
   If cmbTran <> "" Then
      iOldTran = CInt(cmbTran)
   Else
      iOldTran = 1
   End If
   
   cmbTran.Clear
   On Error GoTo diaErr1
   bOnLoad = True
   sSql = "SELECT DISTINCT JITRAN FROM GjitTable WHERE JINAME = '" _
          & Trim(cmbJrn) & "'"
   bSqlRows = GetDataSet(rdoTran, ES_FORWARD)
   If bSqlRows Then
      With rdoTran
         Do Until .EOF
            AddComboStr cmbTran.hWnd, "" & Trim(!JITRAN)
            .MoveNext
         Loop
         .Cancel
      End With
      cmbTran = iOldTran
   Else
      AddComboStr cmbTran.hWnd, "1"
      cmbTran.ListIndex = 0
   End If
   Set rdoTran = Nothing
   bOnLoad = False
   Exit Sub
   diaErr1:
   bOnLoad = False
   Set rdoTran = Nothing
   sProcName = "FillTrans"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Public Sub AddTran()
   Dim bResponse As Byte
   Dim sAccount As String
   Dim sMsg As String
   Dim iTran As Integer
   
   On Error GoTo diaErr1
   MouseCursor 13
   sAccount = Compress(cmbAct)
   iTran = CInt(cmbTran)
   
   On Error Resume Next
   Err = 0
   RdoCon.BeginTrans
   
   sSql = "INSERT INTO GjitTable (JINAME,JITRAN,JIREF,JIDEB," _
          & "JICRD,JIACCOUNT,JIDESC) " _
          & "VALUES('" _
          & cmbJrn & "'," & iTran & "," & Val(txtRef) & "," _
          & CVar(txtDeb) & "," & CVar(txtCrd) & ",'" _
          & sAccount & "','" & Trim(txtcmt) & "')"
   RdoCon.Execute sSql, rdExecDirect
   
   If Err = 0 Then
      On Error GoTo diaErr1
      RdoCon.CommitTrans
      
      ' Update Grid,Trans,Totals
      FillTrans
      FillGrid
      CalcTotals
      MouseCursor 0
   Else
      MouseCursor 0
      On Error GoTo diaErr1
      RdoCon.RollbackTrans
      MsgBox "Couldn't Complete The Transaction.", _
         vbExclamation, Caption
   End If
   
   cmbAct.SetFocus
   Exit Sub
   diaErr1:
   MouseCursor 0
   sProcName = "AddTran"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Public Sub ReviseTran()
   Dim rdoAct As rdoResultset
   
   On Error Resume Next
   Err = 0
   
   sSql = "SELECT GLACCTREF FROM GlacTable WHERE GLACCTNO = '" _
          & Trim(cmbAct) & "'"
   bSqlRows = GetDataSet(rdoAct, ES_FORWARD)
   
   RdoCon.BeginTrans
   sSql = "UPDATE GjitTable SET " _
          & "JIACCOUNT='" & Trim(rdoAct!GLACCTREF) & "'," _
          & "JICRD=" & Val(txtCrd) & "," _
          & "JIDEB=" & Val(txtDeb) & "," _
          & "JIDESC='" & txtcmt & "'" _
          & "WHERE JINAME = '" & Trim(cmbJrn) & "' " _
          & "AND JITRAN = " & CInt(cmbTran) & " " _
          & "AND JIREF = " & CInt(txtRef)
   Set rdoAct = Nothing
   RdoCon.Execute sSql, rdExecDirect
   
   
   If Err = 0 Then
      On Error GoTo diaErr1
      RdoCon.CommitTrans
      FillGrid
      CalcTotals
      MouseCursor 0
   Else
      On Error GoTo diaErr1
      MouseCursor 0
      RdoCon.RollbackTrans
      MsgBox "Couldn't Successfuly Change Journal Entry.", _
         vbExclamation, Caption
   End If
   Exit Sub
   diaErr1:
   sProcName = "ReviseTran"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub cmbact_Click()
   lblActDesc = UpdateActDesc(cmbAct)
End Sub

Private Sub cmbact_LostFocus()
   lblActDesc = UpdateActDesc(cmbAct)
End Sub

Private Sub cmbPost_DropDown()
   ShowCalendar Me
End Sub

Private Sub cmbPost_LostFocus()
   Dim RdoJid As rdoResultset
   cmbPost = CheckDate(cmbPost)
   
   On Error Resume Next
   sSql = "UPDATE GjhdTable SET GJPOST='" & Trim(cmbPost) _
          & "' WHERE GJNAME = '" & Trim(cmbJrn) & "'"
   RdoCon.Execute sSql, rdExecDirect
   If Err = 0 Then
      RdoCon.CommitTrans
   Else
      RdoCon.RollbackTrans
      sProcName = "cmbPost_LostFocus"
      CurrError.Number = Err.Number
      CurrError.Description = Err.Description
      DoModuleErrors Me
   End If
   Set RdoJid = Nothing
End Sub

Private Sub cmbTran_Click()
   SelectJrn
End Sub

Private Sub cmbTran_LostFocus()
   SelectJrn
End Sub

Private Sub cmdDel_Click()
   DelTran
End Sub

Private Sub cmdPost_Click()
   PostJrnl
End Sub

Private Sub SelectJrn()
   Grid1.Enabled = True
   FillGrid
   cmbAct.Enabled = True
   txtDeb.Enabled = True
   txtCrd.Enabled = True
   txtcmt.Enabled = True
   lblActDesc.Enabled = True
   Grid1.Col = 0
   Grid1.Row = 0
   GetGridRow
   CalcTotals
End Sub

Private Sub grid1_Click()
   GetGridRow
End Sub

Private Sub Grid1_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = 46 Then
      DelTran
   End If
End Sub

Private Sub txtCmt_LostFocus()
   txtcmt = CheckLen(txtcmt, 30)
   txtcmt = CheckComments(txtcmt)
End Sub

Private Sub CalcTotals()
   Dim RdoJid As rdoResultset
   Dim rdoTrn As rdoResultset
   
   On Error GoTo diaErr1
   sSql = "SELECT SUM(JICRD) as SumOfJICRD, SUM(JIDEB) as SumOfJIDEB " _
          & "FROM GjitTable WHERE JINAME = '" & Trim(cmbJrn) & "'"
   bSqlRows = GetDataSet(RdoJid, ES_FORWARD)
   
   ' Update Journal Total
   With RdoJid
      If IsNull(!SumOfJIDEB) Then txtDebTotal = 0# Else txtDebTotal = Format(!SumOfJIDEB, "0.00")
      If IsNull(!SumOfJICRD) Then txtCrdTotal = 0# Else txtCrdTotal = Format(!SumOfJICRD, "0.00")
   End With
   
   ' Difference
   txtdiff = Format(Val(txtDebTotal) - Val(txtCrdTotal), "0.00")
   Set RdoJid = Nothing
   
   ' Update Transaction Totals
   sSql = "SELECT SUM(JICRD) as TrnCrd, SUM(JIDEB) as TrnDeb " _
          & "FROM GjitTable WHERE JINAME = '" & Trim(cmbJrn) & "' AND JITRAN = " & CInt(cmbTran)
   bSqlRows = GetDataSet(rdoTrn, ES_FORWARD)
   
   With rdoTrn
      If IsNull(!TrnDeb) Then txtTrnDeb = 0# Else txtTrnDeb = Format(!TrnDeb, "0.00")
      If IsNull(!TrnCrd) Then txtTrnCrd = 0# Else txtTrnCrd = Format(!TrnCrd, "0.00")
   End With
   
   Set rdoTrn = Nothing
   
   If txtCrdTotal <> 0 And txtDebTotal <> 0 Then
      If txtCrdTotal = txtDebTotal Then
         cmdPost.Enabled = True
      Else
         cmdPost.Enabled = False
      End If
   End If
   Exit Sub
   
   diaErr1:
   sProcName = "CalcTotals"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Public Sub GetGridRow()
   ' Grab the info on the row
   Grid1.Row = Grid1.RowSel
   Grid1.Col = 0
   
   ' Did we click on the add new entry row?
   If Trim(Grid1) = "*" Then
      bTrans = 1
      txtRef = MaxRef + 1
      txtDeb = Format(0#, "0.00")
      txtCrd = Format(0#, "0.00")
      'txtcmt = ""
      cmdDel.Enabled = False
   Else
      bTrans = 2
      txtRef = Grid1
      Grid1.Col = 1
      cmbAct.Text = Grid1
      Grid1.Col = 2
      txtDeb = Grid1
      Grid1.Col = 3
      txtCrd = Grid1
      Grid1.Col = 4
      txtcmt = Grid1
      cmdDel.Enabled = True
   End If
   cmbAct.SetFocus
End Sub

Public Sub FillGrid()
   Dim RdoJid As rdoResultset
   Dim sEntry As String
   
   On Error GoTo diaErr1
   MouseCursor 13
   
   sSql = "Select JIDESC,JIACCOUNT,JICRD,JIDEB,JIREF FROM GjitTable " _
          & "WHERE JINAME = '" & Trim(cmbJrn) & "' AND JITRAN = " & CInt(cmbTran) & "ORDER BY JIREF"
   bSqlRows = GetDataSet(RdoJid, ES_FORWARD)
   
   If bSqlRows Then
      Grid1.Clear
      Grid1.Rows = 0
      With RdoJid
         Do Until .EOF
            sEntry = "" & Trim(!JIREF) & Chr(9) & Trim(!JIACCOUNT) & Chr(9) _
                     & "" & CStr(!JIDEB) & Chr(9) & "" & CStr(!JICRD) & Chr(9) & "" & Trim(!JIDESC)
            Grid1.AddItem sEntry ' Add sEntry.
            .MoveNext
         Loop
      End With
      Grid1.Row = 0
   Else
      Grid1.Clear
      Grid1.Rows = 0
   End If
   
   sEntry = "*" & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "(Add New Entry) *"
   Grid1.AddItem sEntry
   MouseCursor 0
   Set RdoJid = Nothing
   
   Exit Sub
   diaErr1:
   Set RdoJid = Nothing
   MouseCursor 0
   sProcName = "FillGrid"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Public Function CheckFiscalYear() As Byte
   Dim RdoFyr As rdoResultset
   On Error GoTo diaErr1
   sSql = "SELECT FYYEAR FROM GlfyTable "
   bSqlRows = GetDataSet(RdoFyr, ES_FORWARD)
   If bSqlRows Then
      CheckFiscalYear = 1
      RdoFyr.Cancel
   Else
      CheckFiscalYear = 0
   End If
   Set RdoFyr = Nothing
   Exit Function
   
   diaErr1:
   sProcName = "checkfisc"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub cmbjrn_Click()
   If Trim(cmbJrn) <> "" Then
      OpenJrn
      SelectJrn
   Else
      MsgBox "Journal Entry Name Cannot Be Blank.", vbInformation
      cmbJrn.SetFocus
   End If
   
End Sub

Private Sub cmbjrn_LostFocus()
   If Not bCancel And Trim(cmbJrn) <> "" Then
      cmbJrn = CheckLen(cmbJrn, 12)
      OpenJrn
      SelectJrn
   ElseIf Not bCancel Then
      MsgBox "Journal Entry Name Cannot Be Blank.", vbInformation
      cmbJrn.SetFocus
   End If
End Sub

Private Sub UnSelect()
   Grid1.Rows = 0
   Grid1.Clear
   Grid1.Enabled = False
   cmbAct.Enabled = False
   txtDeb.Enabled = False
   txtCrd.Enabled = False
   txtcmt.Enabled = False
   lblActDesc.Enabled = False
End Sub

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   bCancel = True
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      SelectHelpTopic Me, "Journal Entry"
      cmdHlp = False
      MouseCursor 0
   End If
End Sub

Private Sub cmdUpdate_Click()
   
   If Val(txtDeb) <> 0 And Val(txtCrd) <> 0 Then
      MsgBox "A Journal Entry Cannot Have Both Debit and Credit Amounts.", _
         vbExclamation, Caption
      txtDeb.SetFocus
      Exit Sub
   End If
   
   If Val(txtDeb) = 0 And Val(txtCrd) = 0 Then
      MsgBox "A Journal Entry Must Have Either A Credit or Debit Amount.", _
         vbExclamation, Caption
      txtDeb.SetFocus
      Exit Sub
   End If
   
   Select Case bTrans
      Case 1
         AddTran
         txtRef = MaxRef + 1
         'txtcrd = Format(0#, "0.00")
         'txtDeb = Format(0#, "0.00")
         cmbAct.SetFocus
      Case 2
         ReviseTran
         bTrans = 0
         Grid1.SetFocus
   End Select
   
   ' Clear out old values buit keep account and comments
   txtDeb = Format(0#, "0.00")
   txtCrd = Format(0#, "0.00")
End Sub

Private Sub Form_Activate()
   Dim sMsg As String
   Dim bResponse As Byte
   
   MdiSect.lblBotPanel = Caption
   
   
   If bOnLoad Then
      'FillCombo
      bGoodYear = CheckFiscalYear()
      
      If bGoodYear Then
         FillJournals
         FillAccounts
      Else
         sMsg = "Fiscal Years Have Not Been Initialized." & vbCr _
                & "Initialize Fiscal Years Now?"
         bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
         
         If bResponse = vbYes Then
            
            Unload Me
            diaGlfyr.Show
         Else
            Unload Me
         End If
         
      End If
      
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   SetDiaPos Me
   FormatControls
   sCurrForm = Caption
   bOnLoad = True
   bCancel = False
   
   Grid1.Rows = 0
   Grid1.Cols = 5
   Grid1.ColWidth(0) = 500
   Grid1.ColWidth(1) = 1500
   Grid1.ColWidth(2) = 1300
   Grid1.ColWidth(3) = 1300
   Grid1.ColWidth(4) = (Grid1.Width - 4600)
   
   cmdPost.Enabled = False
   cmbPost = Format(Now, "mm/dd/yy")
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub


Public Sub FillCombo()
   
   
End Sub


Public Sub FillJournals()
   Dim RdoJrn As rdoResultset
   On Error GoTo diaErr1
   
   cmbJrn.Clear
   sSql = "SELECT GJNAME FROM GjhdTable WHERE GJPOSTED = 0"
   bSqlRows = GetDataSet(RdoJrn, ES_FORWARD)
   If bSqlRows Then
      With RdoJrn
         While Not .EOF
            AddComboStr cmbJrn.hWnd, "" & Trim(!GJNAME)
            .MoveNext
         Wend
      End With
   End If
   Set RdoJrn = Nothing
   Exit Sub
   diaErr1:
   Set RdoJrn = Nothing
   sProcName = "filljournals"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub


Private Sub txtDesc_LostFocus()
   Dim RdoJid As rdoResultset
   
   On Error Resume Next
   
   txtDesc = CheckLen(txtDesc, 30)
   txtDesc = CheckComments(txtDesc)
   
   sSql = "UPDATE GjhdTable SET GJDESC='" & Trim(txtDesc) _
          & "' WHERE GJNAME = '" & Trim(cmbJrn) & "'"
   
   RdoCon.Execute sSql, rdExecDirect
   If Err = 0 Then
      RdoCon.CommitTrans
   Else
      RdoCon.RollbackTrans
      sProcName = "cmbDesc_LostFocus"
      CurrError.Number = Err.Number
      CurrError.Description = Err.Description
      DoModuleErrors Me
   End If
   Set RdoJid = Nothing
End Sub
