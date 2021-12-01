VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form diaGljrn
   BorderStyle = 3 'Fixed Dialog
   Caption = "Journal Entry"
   ClientHeight = 6555
   ClientLeft = 45
   ClientTop = 330
   ClientWidth = 8100
   ClipControls = 0 'False
   ControlBox = 0 'False
   LinkTopic = "Form1"
   MaxButton = 0 'False
   MDIChild = -1 'True
   MinButton = 0 'False
   ScaleHeight = 6555
   ScaleWidth = 8100
   ShowInTaskbar = 0 'False
   Begin VB.TextBox txtDesc
      Height = 285
      Left = 360
      TabIndex = 2
      Tag = "2"
      Top = 960
      Width = 6375
   End
   Begin VB.ComboBox cmbPost
      Height = 315
      Left = 2400
      TabIndex = 1
      Tag = "4"
      Top = 360
      Width = 1335
   End
   Begin VB.TextBox txtcmt
      Height = 285
      Left = 4320
      TabIndex = 8
      Tag = "2"
      Top = 6120
      Width = 3615
   End
   Begin VB.TextBox txtCrd
      Height = 285
      Left = 3240
      TabIndex = 7
      Tag = "1"
      Top = 6120
      Width = 975
   End
   Begin VB.TextBox txtDeb
      Height = 285
      Left = 2160
      TabIndex = 6
      Tag = "1"
      Top = 6120
      Width = 975
   End
   Begin VB.ComboBox cmbAct
      Height = 315
      Left = 720
      TabIndex = 5
      Tag = "3"
      Top = 6120
      Width = 1335
   End
   Begin VB.TextBox txtRef
      BackColor = &H8000000F&
      Enabled = 0 'False
      Height = 285
      Left = 120
      Locked = -1 'True
      TabIndex = 28
      Top = 6120
      Width = 495
   End
   Begin VB.CommandButton cmdUpdate
      Caption = "Update"
      Height = 315
      Left = 7080
      TabIndex = 9
      Top = 5760
      Width = 855
   End
   Begin VB.CommandButton cmdPost
      Caption = "Post"
      Height = 315
      Left = 7080
      TabIndex = 26
      Top = 5160
      Width = 855
   End
   Begin VB.CommandButton cmdAdd
      Caption = "Add"
      Height = 315
      Left = 7080
      TabIndex = 4
      Top = 600
      Width = 855
   End
   Begin VB.TextBox txtCrdTotal
      Alignment = 1 'Right Justify
      BackColor = &H8000000F&
      ForeColor = &H00008000&
      Height = 285
      Left = 3000
      Locked = -1 'True
      TabIndex = 24
      Tag = "1"
      Top = 5160
      Width = 1000
   End
   Begin VB.TextBox txtDebtotal
      Alignment = 1 'Right Justify
      BackColor = &H8000000F&
      ForeColor = &H000000FF&
      Height = 285
      Left = 2040
      Locked = -1 'True
      TabIndex = 23
      Tag = "1"
      Top = 5160
      Width = 1000
   End
   Begin VB.CommandButton cmdDelete
      Caption = "Delete"
      Height = 315
      Left = 7080
      TabIndex = 22
      Top = 1320
      Width = 855
   End
   Begin VB.CommandButton cmdRevise
      Caption = "Revise"
      Height = 315
      Left = 7080
      TabIndex = 20
      Top = 960
      Width = 855
   End
   Begin VB.ComboBox cmbTran
      Height = 315
      Left = 5640
      TabIndex = 3
      Tag = "1"
      Top = 360
      Width = 1095
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1
      Height = 3135
      Left = 120
      TabIndex = 18
      Top = 1920
      Width = 7815
      _ExtentX = 13785
      _ExtentY = 5530
      _Version = 393216
      Rows = 1
      Cols = 5
      FixedRows = 0
      BackColor = -2147483624
      ScrollBars = 2
      SelectionMode = 1
   End
   Begin VB.ComboBox cmbJrn
      Height = 315
      Left = 360
      Sorted = -1 'True
      TabIndex = 0
      Tag = "2"
      ToolTipText = "Select From List"
      Top = 360
      Width = 1890
   End
   Begin VB.CommandButton cmdCan
      Cancel = -1 'True
      Caption = "Close"
      Height = 435
      Left = 7080
      TabIndex = 10
      TabStop = 0 'False
      Top = 90
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
      PictureUp = "diaGljrnRevise.frx":0000
      PictureDn = "diaGljrnRevise.frx":0146
   End
   Begin Threed.SSFrame z2
      Height = 30
      Index = 0
      Left = 0
      TabIndex = 12
      Top = 1440
      Width = 6765
      _Version = 65536
      _ExtentX = 11933
      _ExtentY = 53
      _StockProps = 14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
      Name = "MS Sans Serif"
      Size = 8.26
      Charset = 0
      Weight = 400
      Underline = 0 'False
      Italic = 0 'False
      Strikethrough = 0 'False
      EndProperty
   End
   Begin Threed.SSFrame z2
      Height = 30
      Index = 1
      Left = 120
      TabIndex = 27
      Top = 5640
      Width = 7845
      _Version = 65536
      _ExtentX = 13838
      _ExtentY = 53
      _StockProps = 14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
      Name = "MS Sans Serif"
      Size = 8.26
      Charset = 0
      Weight = 400
      Underline = 0 'False
      Italic = 0 'False
      Strikethrough = 0 'False
      EndProperty
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Description"
      Height = 255
      Index = 0
      Left = 360
      TabIndex = 34
      Top = 720
      Width = 1335
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Ref"
      Height = 255
      Index = 17
      Left = 120
      TabIndex = 33
      Top = 5760
      Width = 615
   End
   Begin VB.Line Line1
      Index = 1
      X1 = 120
      X2 = 6960
      Y1 = 6000
      Y2 = 6000
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Credit Amt"
      Height = 255
      Index = 16
      Left = 3240
      TabIndex = 32
      Top = 5760
      Width = 975
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Comments"
      Height = 255
      Index = 15
      Left = 4320
      TabIndex = 31
      Top = 5760
      Width = 1215
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Debit Amt"
      Height = 255
      Index = 14
      Left = 2160
      TabIndex = 30
      Top = 5760
      Width = 975
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Account"
      Height = 255
      Index = 13
      Left = 720
      TabIndex = 29
      Top = 5760
      Width = 1335
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Post Date"
      Height = 255
      Index = 11
      Left = 2400
      TabIndex = 25
      Top = 120
      Width = 975
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Ref"
      Height = 255
      Index = 8
      Left = 120
      TabIndex = 21
      Top = 1560
      Width = 615
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Transaction"
      Height = 255
      Index = 7
      Left = 5640
      TabIndex = 19
      Top = 120
      Width = 1095
   End
   Begin VB.Line Line1
      Index = 0
      X1 = 120
      X2 = 7920
      Y1 = 1800
      Y2 = 1800
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Credit Amt"
      Height = 255
      Index = 6
      Left = 2880
      TabIndex = 17
      Top = 1560
      Width = 975
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Comments"
      Height = 255
      Index = 5
      Left = 3960
      TabIndex = 16
      Top = 1560
      Width = 1215
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Debit Amt"
      Height = 255
      Index = 4
      Left = 1800
      TabIndex = 15
      Top = 1560
      Width = 975
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Account"
      Height = 255
      Index = 3
      Left = 720
      TabIndex = 14
      Top = 1560
      Width = 975
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Journal ID"
      Height = 255
      Index = 2
      Left = 360
      TabIndex = 13
      Top = 120
      Width = 1335
   End
End
Attribute VB_Name = "diaGljrn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bOnload As Byte ' Prevents calling the same event twice during loads
Dim bCancel As Boolean ' Exit form
Dim bGoodId As Byte
Dim bGoodYear As Byte
Dim iFyear As Integer
Dim iJrnNo As Integer
Dim sJrnl As String ' To tell if the journal has changed
Dim bTrans As Byte ' Are we adding or revising
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
   
   On Error GoTo DiaErr1
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
      MsgBox Trim(cmbJrn) & " Successfully Posted.", _
                  vbInformation, Caption
      
   Else
      RdoCon.RollbackTrans
      MouseCursor 0
      MsgBox "Couldn't Successfuly Post " & Trim(cmbJrn) & ".", _
         vbExclamation, Caption
      
   End If
   Exit Sub
   DiaErr1:
   sProcName = "PostJrnl"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Public Function MaxRef() As Integer
   Dim rdoRef As rdoResultset
   On Error GoTo 0
   ' get next ref #
   sSql = "SELECT Max(JIREF) AS MaxOfJIREF FROM GjitTable WHERE JINAME = '" & Trim(cmbJrn) & " ' AND JITRAN = " & CInt(cmbTran) & ";"
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
   DiaErr1:
   Set rdoRef = Nothing
   sProcName = "MaxRef"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Function

Public Sub FillAccounts()
   ' Fill account combo
   ' Need to add account descriptions
   Dim RdoAct As rdoResultset
   
   sSql = "Qry_FillLowAccounts"
   bSqlRows = GetDataSet(RdoAct, ES_FORWARD)
   
   If bSqlRows Then
      With RdoAct
         Do Until .EOF
            AddComboStr cmbAct.hWnd, "" & Trim(!GLACCTNO)
            .MoveNext
         Loop
      End With
      cmbAct.ListIndex = 0
   End If
   Set RdoAct = Nothing
End Sub

Public Sub OpenJrn()
   Dim RdoJrn As rdoResultset
   Dim iResponse As Integer
   Dim sMsg As String
   Dim iNo As Integer
   Dim iYear As Integer
   ' This function is to be called after the cmbJrn looses focus
   On Error GoTo DiaErr1
   
   sJrnl = cmbJrn
   ' Check if Journal exists
   sSql = "SELECT GJNAME,GJPOST,GJDESC FROM GjhdTable WHERE GJNAME = '" & Trim(sJrnl) & "'"
   bSqlRows = GetDataSet(RdoJrn, ES_FORWARD)
   If bSqlRows Then
      With RdoJrn
         cmbPost = "" & Trim(!GJPOST)
         txtDesc = "" & Trim(!GJDESC)
      End With
      FillTrans
      cmbTran.ListIndex = cmbTran.ListCount - 1
      FillGrid
      CalcTotals
   Else
      ' No Journal exists ask user to make one.
      sMsg = "Journal " & sJrnl & " Does Not Exists.  Do You Wish To Create It?"
      iResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
      If iResponse = vbYes Then
         On Error Resume Next
         ' Get Journal Records
         RdoCon.BeginTrans
         sSql = "INSERT INTO GjhdTable (GJNAME,GJDESC,GJOPEN,GJPOST) VALUES (" _
                & "'" & sJrnl & "'," _
                & "'" & txtDesc & "'," _
                & "'" & Now & "'," _
                & "'" & cmbPost & "')"
         RdoCon.Execute sSql, rdExecDirect
         If Err = 0 Then
            MouseCursor 0
            RdoCon.CommitTrans
            ' Reset the transaction number
            cmbTran.Clear
            cmbTran.AddItem "1"
            cmbTran.ListIndex = 0
            txtDebtotal = 0#
            txtCrdTotal = 0#
            MsgBox sJrnl & " Successfully Opened.", vbInformation, Caption
         Else
            MouseCursor 0
            RdoCon.RollbackTrans
            MsgBox "Couldn't Successfuly Open " & sJrnl & ".", _
               vbExclamation, Caption
         End If
      End If
   End If
   Set RdoJrn = Nothing
   Exit Sub
   DiaErr1:
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
   
   On Error GoTo DiaErr1
   Grid1.Col = 0
   sMsg = "Do You Wish To Delete General Journal Entry Ref " & Grid1 & " ?"
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      MouseCursor 13
      On Error Resume Next
      On Error GoTo 0
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
   DiaErr1:
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
   On Error GoTo DiaErr1
   bOnload = True
   sSql = "SELECT DISTINCT JITRAN FROM GjitTable WHERE JINAME = '" & Trim(cmbJrn) & "'"
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
   bOnload = False
   Exit Sub
   DiaErr1:
   bOnload = False
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
   
   On Error GoTo DiaErr1
   
   sMsg = "You Are About To Enter A New General Journal " & vbCr _
          & "Transaction. To You Wish To Continue?.."
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   
   If bResponse = vbYes Then
      MouseCursor 13
      On Error Resume Next
      Err = 0
      RdoCon.BeginTrans
      sAccount = Compress(cmbAct)
      iTran = CInt(cmbTran)
      sSql = "INSERT INTO GjitTable (JINAME,JITRAN,JIREF,JIDEB," _
             & "JICRD,JIACCOUNT,JIDESC) " _
             & "VALUES('" _
             & cmbJrn & "'," & iTran & "," & Val(txtRef) & "," _
             & Val(txtDeb) & "," & Val(txtCrd) & ",'" _
             & Trim(sAccount) & "','" & Trim(txtcmt) & "')"
      
      RdoCon.Execute sSql, rdExecDirect
      If Err = 0 Then
         On Error GoTo DiaErr1
         RdoCon.CommitTrans
         MsgBox "The Transaction Was Completed Successfully.", _
            vbInformation, Caption
         ' Update Grid,Trans,Totals
         FillTrans
         FillGrid
         CalcTotals
         MouseCursor 0
      Else
         MouseCursor 0
         On Error GoTo DiaErr1
         RdoCon.RollbackTrans
         MsgBox "Couldn't Complete The Transaction.", _
            vbExclamation, Caption
      End If
   Else
      cmbAct.SetFocus
   End If
   Exit Sub
   DiaErr1:
   MouseCursor 0
   sProcName = "AddTran"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Public Sub ReviseTran()
   On Error Resume Next
   Err = 0
   RdoCon.BeginTrans
   sSql = "UPDATE GjitTable SET " _
          & "JIACCOUNT='" & Trim(cmbAct) & "'," _
          & "JICRD=" & Val(txtCrd) & "," _
          & "JIDEB=" & Val(txtDeb) & "," _
          & "JIDESC='" & txtcmt & "'" _
          & "WHERE JINAME = '" & Trim(cmbJrn) & "' " _
          & "AND JITRAN = " & CInt(cmbTran) & " " _
          & "AND JIREF = " & CInt(txtRef)
   RdoCon.Execute sSql, rdExecDirect
   
   If Err = 0 Then
      On Error GoTo DiaErr1
      RdoCon.CommitTrans
      FillGrid
      CalcTotals
      MouseCursor 0
   Else
      On Error GoTo DiaErr1
      MouseCursor 0
      RdoCon.RollbackTrans
      MsgBox "Couldn't Successfuly Update Journal Entry.", _
         vbExclamation, Caption
   End If
   Exit Sub
   DiaErr1:
   sProcName = "ReviseTran"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub cmbJrn_GotFocus()
   TurnEntry False
End Sub

Private Sub cmbPost_DropDown()
   ShowCalendar Me
End Sub

Private Sub cmbPost_GotFocus()
   TurnEntry False
End Sub

Private Sub cmbPost_LostFocus()
   Dim RdoJid As rdoResultset
   cmbPost = CheckDate(cmbPost)
   ' Update the posting date
   On Error Resume Next
   sSql = "UPDATE GjhdTable SET GJPOST='" & Trim(cmbPost) & "' WHERE GJNAME = '" & Trim(cmbJrn) & "'"
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

Private Sub cmdDelete_Click()
   DelTran
End Sub

Private Sub cmdPost_Click()
   PostJrnl
End Sub

Private Sub Grid1_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = 46 Then
      DelTran
   End If
End Sub



Private Sub txtCmt_LostFocus()
   txtcmt = CheckLen(txtcmt, 30)
End Sub

Private Sub txtCrd_LostFocus()
   txtCrd = Format(txtCrd, "#######0.00")
End Sub

Private Sub txtDeb_LostFocus()
   txtDeb = Format(txtDeb, "#######0.00")
End Sub

Public Sub TurnEntry(bOn As Boolean)
   ' Turn on or off the trans dataentry fields
   txtRef.Enabled = bOn
   cmbAct.Enabled = bOn
   txtDeb.Enabled = bOn
   txtCrd.Enabled = bOn
   txtcmt.Enabled = bOn
   cmdUpdate.Enabled = bOn
   
   z1(13).Enabled = bOn
   z1(14).Enabled = bOn
   z1(15).Enabled = bOn
   z1(16).Enabled = bOn
   z1(17).Enabled = bOn
   
   FillAccounts
   If Not bOn Or bTrans = 1 Then
      txtRef = ""
      cmbAct.ListIndex = 0
      txtDeb = ""
      txtCrd = ""
      txtcmt = ""
      cmdAdd.Enabled = True
   End If
End Sub

Public Sub CalcTotals()
   Dim RdoJid As rdoResultset
   On Error GoTo DiaErr1
   sSql = "SELECT SUM(JICRD) as SumOfJICRD, SUM(JIDEB) as SumOfJIDEB " _
          & "FROM GjitTable WHERE JINAME = '" & Trim(cmbJrn) & "'"
   bSqlRows = GetDataSet(RdoJid, ES_FORWARD)
   
   With RdoJid
      If IsNull(!SumOfJIDEB) Then
         txtDebtotal = 0#
      Else
         txtDebtotal = Format(!SumOfJIDEB, "#######0.00")
      End If
      
      If IsNull(!SumOfJIDEB) Then
         txtCrdTotal = 0#
      Else
         txtCrdTotal = Format(!SumOfJICRD, "#######0.00")
      End If
      .Cancel
   End With
   
   If txtCrdTotal = txtDebtotal And txtCrdTotal <> 0 And txtDebtotal <> 0 Then
      cmdPost.Enabled = True
   Else
      cmdPost.Enabled = False
   End If
   
   Exit Sub
   DiaErr1:
   sProcName = "CalcTotals"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Public Sub GetGridRow()
   'grab the info on the row
   cmdRevise.Enabled = False
   cmdDelete.Enabled = False
   bTrans = 2
   TurnEntry True
   Grid1.Row = Grid1.RowSel
   
   Grid1.Col = 0
   txtRef = Grid1
   Grid1.Col = 1
   cmbAct.Text = Grid1
   
   Grid1.Col = 2
   txtDeb = Grid1
   Grid1.Col = 3
   txtCrd = Grid1
   Grid1.Col = 4
   txtcmt = Grid1
   
   
   cmbAct.SetFocus
End Sub

Public Sub FillGrid()
   Dim RdoJid As rdoResultset
   Dim sEntry As String
   
   On Error GoTo DiaErr1
   MouseCursor 13
   
   sSql = "Select JIDESC,JIACCOUNT,JICRD,JIDEB,JIREF FROM GjitTable WHERE JINAME = '" & cmbJrn & "' AND JITRAN = " & CInt(cmbTran)
   bSqlRows = GetDataSet(RdoJid, ES_FORWARD)
   If bSqlRows Then
      Grid1.Clear
      Grid1.Rows = 0
      With RdoJid
         Do Until .EOF
            sEntry = "" & Trim(!JIREF) & Chr(9) & Trim(!JIACCOUNT) & Chr(9) _
                     & "" & Format(!JIDEB, "######0.00") & Chr(9) & "" & Format(!JICRD, "######0.00") & Chr(9) & "" & Trim(!JIDESC)
            Grid1.AddItem sEntry ' Add sEntry.
            .MoveNext
         Loop
      End With
      Grid1.Row = 0
   Else
      Grid1.Clear
      Grid1.Rows = 0
   End If
   
   MouseCursor 0
   Set RdoJid = Nothing
   cmdAdd.Enabled = True
   Exit Sub
   DiaErr1:
   Set RdoJid = Nothing
   MouseCursor 0
   sProcName = "FillGrid"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Public Function CheckFiscalYear() As Byte
   Dim RdoFyr As rdoResultset
   On Error GoTo DiaErr1
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
   
   DiaErr1:
   sProcName = "checkfisc"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub cmbjrn_Click()
   TurnEntry False
   
   ' Make sure we have a new journal before opening
   If cmbJrn <> sJrnl Then
      OpenJrn
      'FillJournals
   End If
End Sub

Private Sub cmbjrn_LostFocus()
   cmbJrn = CheckLen(cmbJrn, 12)
   
   If Not bCancel Then
      If Len(Trim(cmbJrn)) Then
         bGoodId = True
         OpenJrn
      Else
         bGoodId = False
      End If
   End If
End Sub


Private Sub cmbTran_Click()
   If Len(cmbTran) > 0 And Not bOnload Then
      FillGrid
   End If
End Sub

Private Sub cmbTran_GotFocus()
   TurnEntry False
End Sub

Private Sub cmbTran_LostFocus()
   'If bCancel Then Exit Sub
   If Len(cmbTran) > 0 And Not bOnload Then
      FillGrid
   End If
   If cmbTran = "" Then
      cmbTran.ListIndex = 0
   End If
End Sub

Private Sub cmdAdd_Click()
   bTrans = 1 'Add
   TurnEntry True
   cmdAdd.Enabled = False
   cmdDelete.Enabled = False
   cmdRevise.Enabled = False
   txtRef = CInt(MaxRef + 1)
   txtCrd = Format("0", "#######0.00")
   txtDeb = Format("0", "#######0.00")
   cmbAct.SetFocus
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

Private Sub cmdRevise_Click()
   GetGridRow
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
         txtCrd = Format("0", "#######0.00")
         txtDeb = Format("0", "#######0.00")
         cmbAct.SetFocus
      Case 2
         ReviseTran
         bTrans = 0
         Grid1.SetFocus
   End Select
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnload Then
      FillCombo
      bOnload = False
      cmbJrn.SetFocus
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Click()
   TurnEntry False
   cmdDelete.Enabled = True
   cmdRevise.Enabled = True
End Sub

Private Sub Form_Load()
   SetDiaPos Me
   FormatControls
   sCurrForm = Caption
   bOnload = True
   bCancel = False
   Grid1.ColWidth(0) = 500
   Grid1.ColWidth(1) = 1100
   Grid1.ColWidth(2) = 1100
   Grid1.ColWidth(3) = 1100
   Grid1.ColWidth(4) = 4000
   TurnEntry False
   ' Set the buttons
   cmdPost.Enabled = False
   cmdDelete.Enabled = False
   cmdAdd.Enabled = False
   cmdRevise.Enabled = False
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
   Dim sMsg As String
   Dim bResponse As Byte
   
   bGoodYear = CheckFiscalYear()
   If bGoodYear Then
      FillJournals
   Else
      sMsg = "Fiscal Years Have Not Been Initialized." & vbCr _
             & "Initialize Fiscal Years Now?"
      bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
      If bResponse = vbYes Then
         diaGlfyr.Show
         Unload Me
      Else
         Unload Me
      End If
   End If
   
End Sub


Public Sub FillJournals()
   Dim RdoJrn As rdoResultset
   On Error GoTo DiaErr1
   On Error GoTo 0
   cmbJrn.Clear
   
   sSql = "SELECT GJNAME FROM GjhdTable WHERE GJPOSTED = 0"
   bSqlRows = GetDataSet(RdoJrn, ES_FORWARD)
   If bSqlRows Then
      With RdoJrn
         Do Until .EOF
            AddComboStr cmbJrn.hWnd, "" & Trim(!GJNAME)
            .MoveNext
         Loop
         .Cancel
      End With
      cmbJrn.ListIndex = 0
   End If
   Set RdoJrn = Nothing
   Exit Sub
   DiaErr1:
   Set RdoJrn = Nothing
   sProcName = "filljournals"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub


Private Sub Grid1_DblClick()
   GetGridRow
End Sub

Private Sub grid1_GotFocus()
   TurnEntry False
   cmdRevise.Enabled = True
   cmdDelete.Enabled = True
End Sub

Private Sub Grid1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      GetGridRow
   End If
End Sub

Private Sub txtDesc_GotFocus()
   TurnEntry False
End Sub

Private Sub txtDesc_LostFocus()
   Dim RdoJid As rdoResultset
   ' Update the posting date
   On Error Resume Next
   txtDesc = CheckLen(txtDesc, 30)
   sSql = "UPDATE GjhdTable SET GJDESC='" & Trim(txtDesc) & "' WHERE GJNAME = '" & Trim(cmbJrn) & "'"
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
