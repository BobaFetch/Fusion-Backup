VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaCdivs
   BorderStyle = 3 'Fixed Dialog
   Caption = "Divisions"
   ClientHeight = 2895
   ClientLeft = 1200
   ClientTop = 855
   ClientWidth = 5700
   ClipControls = 0 'False
   ControlBox = 0 'False
   ForeColor = &H8000000F&
   LinkTopic = "Form1"
   MaxButton = 0 'False
   MDIChild = -1 'True
   MinButton = 0 'False
   ScaleHeight = 2895
   ScaleWidth = 5700
   ShowInTaskbar = 0 'False
   Begin ResizeLibCtl.ReSize ReSize1
      Left = 4800
      Top = 2640
      _Version = 196615
      _ExtentX = 741
      _ExtentY = 741
      _StockProps = 0
      Enabled = -1 'True
      FormMinWidth = 0
      FormMinHeight = 0
      FormDesignHeight = 2895
      FormDesignWidth = 5700
   End
   Begin VB.TextBox txtDsc
      Height = 285
      Left = 1360
      TabIndex = 1
      Top = 960
      Width = 3495
   End
   Begin VB.CheckBox optUse
      Alignment = 1 'Right Justify
      Caption = "Force Divisions"
      Enabled = 0 'False
      Height = 315
      Left = 120
      TabIndex = 3
      Top = 2400
      Value = 1 'Checked
      Width = 1455
   End
   Begin VB.TextBox txtCmt
      Height = 975
      Left = 1360
      MultiLine = -1 'True
      TabIndex = 2
      Top = 1320
      Width = 3475
   End
   Begin VB.ComboBox cmbDiv
      Height = 315
      Left = 1360
      Sorted = -1 'True
      TabIndex = 0
      Tag = "3"
      ToolTipText = "Enter/Revise A Division (2 char)"
      Top = 600
      Width = 660
   End
   Begin VB.CommandButton cmdCan
      Cancel = -1 'True
      Caption = "Close"
      Height = 435
      Left = 4680
      TabIndex = 6
      TabStop = 0 'False
      Top = 90
      Width = 875
   End
   Begin Threed.SSRibbon cmdHlp
      Height = 225
      Left = 0
      TabIndex = 9
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
      PictureUp = "diaCdivs.frx":0000
      PictureDn = "diaCdivs.frx":0146
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "(Forces Use Of Fixed Divisions In All Sections)"
      Height = 255
      Index = 3
      Left = 1680
      TabIndex = 8
      Top = 2445
      Width = 3975
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Comment"
      Height = 255
      Index = 2
      Left = 120
      TabIndex = 7
      Top = 1320
      Width = 1215
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Description"
      Height = 255
      Index = 1
      Left = 120
      TabIndex = 5
      Top = 960
      Width = 1215
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Division"
      Height = 255
      Index = 0
      Left = 120
      TabIndex = 4
      Top = 600
      Width = 1215
   End
End
Attribute VB_Name = "diaCdivs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001, ES/2002) is the property of                     ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***

'See the UpdateTables prodecure for database revisions

Option Explicit

Dim RdoQry As rdoQuery

Dim bOnLoad As Byte
Dim bGoodDiv As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd



Private Sub cmbDiv_Click()
   bGoodDiv = GetDivision()
   
End Sub


Private Sub cmbDiv_LostFocus()
   cmbDiv = CheckLen(cmbDiv, 2)
   If Len(cmbDiv) Then
      cmbDiv = Compress(cmbDiv)
      bGoodDiv = GetDivision()
      If Not bGoodDiv Then AddDivision
   Else
      bGoodDiv = False
   End If
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   cmbDiv = ""
   
End Sub


Private Sub cmdHlp_Click(Value As Integer)
   Dim l&
   If cmdHlp Then
      MouseCursor 13
      l& = WinHelp(hWnd, sReportPath & "EsiAdmn.hlp", HELP_KEY, "Divisions")
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub


Private Sub Form_Activate()
   If bOnLoad Then
      FillDivisions Me
      bOnLoad = False
      If cmbDiv.ListCount > 0 Then
         cmbDiv = cmbDiv.List(0)
         bGoodDiv = GetDivision()
      End If
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   SetDiaPos Me
   FormatControls
   sCurrForm = Caption
   sSql = "SELECT DIVREF,DIVDESC,DIVCOMT FROM " _
          & "CdivTable WHERE DIVREF= ? "
   Set RdoQry = RdoCon.CreateQuery("", sSql)
   RdoQry.MaxRows = 1
   Dim sString$
   bOnLoad = True
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set diaCdivs = Nothing
   
End Sub

Public Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
   
   
End Sub

Private Sub optUse_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub txtCmt_LostFocus()
   txtCmt = CheckLen(txtCmt, 255)
   txtCmt = StrCase(txtCmt, ES_FIRSTWORD)
   On Error Resume Next
   If bGoodDiv Then
      rdoRes.Edit
      rdoRes!DIVCOMT = "" & txtCmt
      rdoRes.Update
      If Err > 0 Then ValidateEdit Me
   End If
   
End Sub


Private Sub txtDsc_LostFocus()
   txtDsc = CheckLen(txtDsc, 40)
   txtDsc = StrCase(txtDsc, vbProperCase)
   On Error Resume Next
   If bGoodDiv Then
      rdoRes.Edit
      rdoRes!DIVDESC = "" & txtDsc
      rdoRes.Update
      If Err > 0 Then ValidateEdit Me
   End If
   
End Sub



Public Function GetDivision() As Byte
   Dim sDivision As String
   sDivision = cmbDiv
   
   On Error GoTo DiaErr1
   RdoQry(0) = sDivision
   bSqlRows = GetQuerySet(rdoRes, RdoQry, ES_KEYSET)
   If bSqlRows Then
      With rdoRes
         cmbDiv = "" & Trim(!DIVREF)
         txtDsc = "" & Trim(!DIVDESC)
         txtCmt = "" & Trim(!DIVCOMT)
         .Cancel
      End With
      GetDivision = True
   Else
      txtDsc = ""
      txtCmt = ""
      GetDivision = False
   End If
   Exit Function
   
   DiaErr1:
   sProcName = "getdivisi"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Public Sub AddDivision()
   Dim bResponse As Byte
   Dim sMsg As String
   Dim sDivision As String
   
   sDivision = cmbDiv
   sMsg = sDivision & " Wasn't Found. Add The Division?"
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      On Error GoTo DiaErr1
      sSql = "INSERT INTO CdivTable (DIVREF) " _
             & "VALUES('" & sDivision & "')"
      RdoCon.Execute sSql, rdExecDirect
      If RdoCon.RowsAffected Then
         Sysmsg "Division Added.", True
         cmbDiv = sDivision
         cmbDiv.AddItem sDivision
         bGoodDiv = GetDivision()
         On Error Resume Next
         txtDsc.SetFocus
      Else
         MsgBox "Couldn't The Add Division.", vbExclamation, Caption
      End If
   Else
      CancelTrans
   End If
   Exit Sub
   
   DiaErr1:
   sProcName = "adddivisi"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub
