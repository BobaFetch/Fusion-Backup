VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form diaJRf02a
   BorderStyle = 3 'Fixed Dialog
   Caption = "Close Journals"
   ClientHeight = 5100
   ClientLeft = 45
   ClientTop = 330
   ClientWidth = 6465
   ClipControls = 0 'False
   ControlBox = 0 'False
   LinkTopic = "Form1"
   MaxButton = 0 'False
   MDIChild = -1 'True
   MinButton = 0 'False
   ScaleHeight = 5100
   ScaleWidth = 6465
   ShowInTaskbar = 0 'False
   Begin VB.ComboBox cmbTyp
      ForeColor = &H00800000&
      Height = 315
      Left = 1200
      TabIndex = 3
      Tag = "8"
      ToolTipText = "Select Journal Type From List"
      Top = 360
      Width = 3135
   End
   Begin VB.CommandButton cmdCls
      Caption = "&Close Jrn"
      Enabled = 0 'False
      Height = 315
      Left = 5520
      TabIndex = 2
      TabStop = 0 'False
      ToolTipText = "Close This Journal"
      Top = 720
      Width = 875
   End
   Begin VB.CommandButton cmdCan
      Cancel = -1 'True
      Caption = "Close"
      Height = 435
      Left = 5520
      TabIndex = 0
      TabStop = 0 'False
      Top = 120
      Width = 875
   End
   Begin Threed.SSRibbon cmdHlp
      Height = 225
      Left = 0
      TabIndex = 1
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
      PictureUp = "diaJrnpj.frx":0000
      PictureDn = "diaJrnpj.frx":0146
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1
      Height = 3495
      Left = 120
      TabIndex = 7
      ToolTipText = "Open Journals Of Selected Type"
      Top = 1200
      Width = 6255
      _ExtentX = 11033
      _ExtentY = 6165
      _Version = 393216
      Cols = 5
      FixedRows = 0
      FixedCols = 0
      WordWrap = -1 'True
      HighLight = 2
      GridLines = 0
      ScrollBars = 2
      SelectionMode = 1
   End
   Begin ResizeLibCtl.ReSize ReSize2
      Left = 4920
      Top = 0
      _Version = 196615
      _ExtentX = 741
      _ExtentY = 741
      _StockProps = 0
      Enabled = -1 'True
      FormMinWidth = 0
      FormMinHeight = 0
      FormDesignHeight = 5100
      FormDesignWidth = 6465
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Journal            "
      BeginProperty Font
      Name = "MS Sans Serif"
      Size = 8.25
      Charset = 0
      Weight = 400
      Underline = -1 'True
      Italic = 0 'False
      Strikethrough = 0 'False
      EndProperty
      Height = 255
      Index = 4
      Left = 120
      TabIndex = 12
      Top = 960
      Width = 1095
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Start            "
      BeginProperty Font
      Name = "MS Sans Serif"
      Size = 8.25
      Charset = 0
      Weight = 400
      Underline = -1 'True
      Italic = 0 'False
      Strikethrough = 0 'False
      EndProperty
      Height = 255
      Index = 2
      Left = 1320
      TabIndex = 11
      Top = 960
      Width = 1095
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Description                                                                        "
      BeginProperty Font
      Name = "MS Sans Serif"
      Size = 8.25
      Charset = 0
      Weight = 400
      Underline = -1 'True
      Italic = 0 'False
      Strikethrough = 0 'False
      EndProperty
      Height = 255
      Index = 3
      Left = 3240
      TabIndex = 10
      Top = 960
      Width = 3135
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "End             "
      BeginProperty Font
      Name = "MS Sans Serif"
      Size = 8.25
      Charset = 0
      Weight = 400
      Underline = -1 'True
      Italic = 0 'False
      Strikethrough = 0 'False
      EndProperty
      Height = 255
      Index = 6
      Left = 2280
      TabIndex = 9
      Top = 960
      Width = 1095
   End
   Begin VB.Label lblStatus
      Caption = "Selected Journal:"
      Height = 255
      Left = 120
      TabIndex = 8
      Top = 4800
      Width = 6255
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Description"
      Height = 255
      Index = 1
      Left = 4080
      TabIndex = 6
      Top = 3840
      Width = 15
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Journal Type"
      Height = 255
      Index = 0
      Left = 120
      TabIndex = 5
      Top = 360
      Width = 1215
   End
   Begin VB.Label lblTyp
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      Height = 285
      Left = 4560
      TabIndex = 4
      Top = 360
      Width = 375
   End
End
Attribute VB_Name = "diaJRf02a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001, ES/2002) is the property of                     ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***

'See the UpdateTables prodecure for database revisions

'************************************************************************************
' diaJRf02a - Close Journal
'
'
' Created: (nth)
' Revions:
'
'************************************************************************************

Option Explicit

Dim bOnLoad As Byte
Dim bGoodYear As Byte
Dim bThisYear As Byte
Dim iNextJournal As Integer

Dim sJournals(13, 2) As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'***********************************************************************************

Private Sub CloseJrn()
   Dim bResponse As Byte
   Dim smsg As String
   Dim sJrn As String
   
   On Error GoTo DiaErr1
   
   sJrn = Grid1.Text
   
   smsg = "Are You Certain That You Wish To" & vbCr _
          & "Close " & sJrn & " ?"
   bResponse = MsgBox(smsg, ES_NOQUESTION, Caption)
   
   If bResponse = vbYes Then
      sSql = "UPDATE JrhdTable SET MJCLOSED='" & Format(Now, "mm/dd/yyyy") & "' " _
             & "WHERE MJGLJRNL ='" & sJrn & "'"
      RdoCon.Execute sSql
      
      If Err = 0 Then
         RdoCon.CommitTrans
         Sysmsg sJrn & " Successful Closed.", 1
         FillJournals
      Else
         RdoCon.RollbackTrans
         MsgBox "Could Not Close Journal." & vbCrLf & "Transaction Canceled.", _
            vbExclamation, Caption
      End If
   End If
   Exit Sub
   
   DiaErr1:
   sProcName = "CloseJrnl"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub cmbTyp_Click()
   lblTyp = sJournals(cmbTyp.ListIndex, 1)
   FillJournals
End Sub

Private Sub cmbTyp_LostFocus()
   Dim b As Byte
   Dim i As Integer
   For i = 0 To cmbTyp.ListCount - 1
      If cmbTyp = cmbTyp.List(i) Then b = True
   Next
   If Not b Then
      Beep
      cmbTyp = cmbTyp.List(0)
      lblTyp = sJournals(0, 1)
   End If
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
   SetDiaPos Me
   FormatControls
   sJournals(0, 0) = "Sales Journal"
   sJournals(0, 1) = "SJ"
   sJournals(1, 0) = "Purchases Journal"
   sJournals(1, 1) = "PJ"
   sJournals(2, 0) = "Cash Receipts Journal"
   sJournals(2, 1) = "CR"
   sJournals(3, 0) = "Cash Disbursements-Computer Checks"
   sJournals(3, 1) = "CC"
   sJournals(4, 0) = "Cash Disbursements-External Checks"
   sJournals(4, 1) = "XC"
   sJournals(5, 0) = "Payroll Labor Journal"
   sJournals(5, 1) = "PL"
   sJournals(6, 0) = "Payroll Disbursements Journal"
   sJournals(6, 1) = "PD"
   sJournals(7, 0) = "Time Journal"
   sJournals(7, 1) = "TJ"
   sJournals(8, 0) = "Inventory Journal"
   sJournals(8, 1) = "IJ"
   
   sCurrForm = Caption
   bOnLoad = True
   
   Grid1.Cols = 4
   Grid1.ColWidth(0) = 1100
   Grid1.ColWidth(1) = 1000
   Grid1.ColWidth(2) = 1000
   Grid1.ColWidth(3) = 5000
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
   Dim smsg As String
   Dim i As Integer
   On Error GoTo DiaErr1
   For i = 0 To 8
      AddComboStr cmbTyp.hWnd, sJournals(i, 0)
   Next
   cmbTyp = cmbTyp.List(0)
   lblTyp = sJournals(0, 1)
   bGoodYear = CheckFiscalYear()
   If bGoodYear Then
      FillJournals
   Else
      smsg = "Fiscal Years Have Not Been Initialized." & vbCr _
             & "Initialize Fiscal Years Now?"
      bResponse = MsgBox(smsg, ES_YESQUESTION, Caption)
      If bResponse = vbYes Then
         diaGLe03a.Show
         Unload Me
      Else
         
         cmdCls.Enabled = False
      End If
   End If
   
   Exit Sub
   
   DiaErr1:
   sProcName = "fillcombo"
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


Public Sub FillJournals()
   Dim sItem As String
   Dim RdoLst As rdoResultset
   On Error GoTo DiaErr1
   
   Grid1.Clear
   Grid1.Rows = 0
   
   sSql = "SELECT MJGLJRNL,MJSTART,MJEND,MJDESCRIPTION " _
          & "FROM JrhdTable WHERE MJTYPE='" & lblTyp & "' AND (MJCLOSED IS NULL) " _
          
   bSqlRows = GetDataSet(RdoLst, ES_FORWARD)
   If bSqlRows Then
      cmdCls.Enabled = True
      With RdoLst
         While Not .EOF
            sItem = !MJGLJRNL & Chr(9) _
                    & Format(!MJSTART, "mm/dd/yyyy") & Chr(9) _
                    & Format(!MJEND, "mm/dd/yyyy") & Chr(9) _
                    & Trim(!MJDESCRIPTION)
            Grid1.AddItem sItem
            .MoveNext
         Wend
         .Cancel
      End With
   Else
      cmdCls.Enabled = False
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
   lblStatus = "Selected Journal: " & Grid1.Text
End Sub


Private Sub Grid1_EnterCell()
   lblStatus = "Selected Journal: " & Grid1.Text
End Sub

Private Sub Grid1_GotFocus()
   lblStatus = "Selected Journal: " & Grid1.Text
End Sub
