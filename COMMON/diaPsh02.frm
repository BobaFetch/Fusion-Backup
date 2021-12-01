VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Begin VB.Form diaPsh02
   BorderStyle = 3 'Fixed Dialog
   Caption = "Manufacturing Order History"
   ClientHeight = 2520
   ClientLeft = 45
   ClientTop = 330
   ClientWidth = 7260
   ClipControls = 0 'False
   ControlBox = 0 'False
   ForeColor = &H8000000F&
   LinkTopic = "Form1"
   LockControls = -1 'True
   MaxButton = 0 'False
   MDIChild = -1 'True
   MinButton = 0 'False
   ScaleHeight = 2520
   ScaleWidth = 7260
   ShowInTaskbar = 0 'False
   Begin VB.ComboBox txtBeg
      Height = 315
      Left = 1920
      TabIndex = 1
      Tag = "4"
      Top = 1800
      Width = 1095
   End
   Begin VB.ComboBox txtEnd
      Height = 315
      Left = 4080
      TabIndex = 2
      Tag = "4"
      Top = 1800
      Width = 1095
   End
   Begin VB.Frame fraPrn
      BorderStyle = 0 'None
      Height = 495
      Left = 6120
      TabIndex = 12
      Top = 450
      Width = 1095
      Begin VB.CommandButton optDis
         Height = 330
         Left = 0
         Picture = "diaPsh02.frx":0000
         Style = 1 'Graphical
         TabIndex = 3
         ToolTipText = "Display The Report"
         Top = 120
         UseMaskColor = -1 'True
         Width = 495
      End
      Begin VB.CommandButton optPrn
         Height = 330
         Left = 560
         Picture = "diaPsh02.frx":017E
         Style = 1 'Graphical
         TabIndex = 4
         ToolTipText = "Print The Report"
         Top = 120
         UseMaskColor = -1 'True
         Width = 495
      End
   End
   Begin VB.CommandButton cmdCan
      Cancel = -1 'True
      Caption = "Close"
      Height = 360
      Left = 6120
      TabIndex = 11
      TabStop = 0 'False
      Top = 90
      Width = 1065
   End
   Begin VB.ComboBox cmbPrt
      Height = 315
      Left = 1920
      Sorted = -1 'True
      TabIndex = 0
      Tag = "3"
      ToolTipText = "Contains Part Numbers With Manufacturing Orders"
      Top = 960
      Width = 3545
   End
   Begin Threed.SSRibbon cmdHlp
      Height = 225
      Left = 0
      TabIndex = 5
      ToolTipText = "Subject Help"
      Top = 0
      Width = 255
      _Version = 65536
      _ExtentX = 450
      _ExtentY = 397
      _StockProps = 65
      BackColor = 12632256
      GroupAllowAllUp = -1 'True
      Autosize = 2
      RoundedCorners = 0 'False
      BevelWidth = 0
      Outline = 0 'False
      PictureUp = "diaPsh02.frx":0308
      PictureDn = "diaPsh02.frx":044E
   End
   Begin ResizeLibCtl.ReSize ReSize1
      Left = 6720
      Top = 2280
      _Version = 196615
      _ExtentX = 741
      _ExtentY = 741
      _StockProps = 0
      Enabled = -1 'True
      FormMinWidth = 0
      FormMinHeight = 0
      FormDesignHeight = 2520
      FormDesignWidth = 7260
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Runs Scheduled"
      Height = 255
      Index = 4
      Left = 360
      TabIndex = 15
      Top = 1560
      Width = 1935
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Type"
      Height = 255
      Index = 15
      Left = 5520
      TabIndex = 14
      Top = 1320
      Width = 615
   End
   Begin VB.Label lblTyp
      BorderStyle = 1 'Fixed Single
      Height = 285
      Left = 6120
      TabIndex = 13
      Top = 1320
      Width = 300
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "(Blank For All)"
      Height = 255
      Index = 3
      Left = 5520
      TabIndex = 10
      Top = 1800
      Width = 3255
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Through"
      Height = 255
      Index = 2
      Left = 3120
      TabIndex = 9
      Top = 1800
      Width = 1815
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Complete From:"
      Height = 255
      Index = 1
      Left = 360
      TabIndex = 8
      Top = 1800
      Width = 1935
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Part Number"
      Height = 255
      Index = 0
      Left = 360
      TabIndex = 7
      Top = 960
      Width = 1095
   End
   Begin VB.Label lblDsc
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      Height = 285
      Left = 1920
      TabIndex = 6
      Top = 1320
      Width = 3255
   End
End
Attribute VB_Name = "diaPsh02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of            ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
Option Explicit
Dim bGoodPart As Byte
Dim bOnLoad As Byte

Dim sPartNumber As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtEnd = Format(ES_SYSDATE, "mm/dd/yy")
   txtBeg = "01/01/" & Right(txtEnd, 2)
   
End Sub

Private Sub cmbPrt_Click()
   bGoodPart = GetPart()
   
End Sub

Private Sub cmbPrt_LostFocus()
   cmbPrt = CheckLen(cmbPrt, 30)
   bGoodPart = GetPart()
   
End Sub

Private Sub PrintReport()
   Dim sBegDate As String
   Dim sEndDate As String
   
   
   If Len(Trim(txtBeg)) = 0 Then txtBeg = "ALL"
   If Len(Trim(txtEnd)) = 0 Then txtEnd = "ALL"
   If txtBeg = "ALL" Then
      sBegDate = "1995,01,01"
   Else
      sBegDate = Format(txtBeg, "yyyy,mm,dd")
   End If
   If txtEnd = "ALL" Then
      sEndDate = "2024,12,31"
   Else
      sEndDate = Format(txtEnd, "yyyy,mm,dd")
   End If
   MouseCursor 13
   On Error GoTo Psh02
   SetMdiReportsize MdiSect
   sPartNumber = Compress(cmbPrt)
   MdiSect.Crw.Formulas(0) = "CompanyName='" & sFacility & "'"
   MdiSect.Crw.Formulas(1) = "Includes='" & txtBeg & "'"
   MdiSect.Crw.Formulas(2) = "Includes2='" & txtEnd & "'"
   sCustomReport = GetCustomReport("prdsh02")
   MdiSect.Crw.ReportFileName = sReportPath & sCustomReport
   sSql = "{RunsTable.RUNREF} = '" & sPartNumber & "' AND " _
          & "{RunsTable.RUNSCHED} in Date(" & Format(sBegDate, "yyyy,mm,dd") & ") " _
          & "to Date(" & Format(sEndDate, "yyyy,mm,dd") & ")"
   MdiSect.Crw.SelectionFormula = sSql
   SetCrystalAction Me
   MouseCursor 0
   Exit Sub
   
   Psh02:
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   Resume Psh02a
   Psh02a:
   DoModuleErrors Me
   
End Sub




Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      OpenWebHelp "hs907"
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub



Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillAllRuns cmbPrt
      bGoodPart = GetPart()
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   SetDiaPos Me
   FormatControls
   
   bOnLoad = 1
   Show
   
End Sub




Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set diaPsh02 = Nothing
   
End Sub





Private Function GetPart() As Byte
   Dim rdoPrt As rdoResultset
   sPartNumber = Compress(cmbPrt)
   On Error GoTo DiaErr1
   If Len(sPartNumber) > 0 Then
      sSql = "SELECT PARTREF,PARTNUM,PADESC,PALEVEL FROM PartTable WHERE PARTREF='" & sPartNumber & "'"
      bSqlRows = GetDataSet(rdoPrt, ES_FORWARD)
      If bSqlRows Then
         With rdoPrt
            cmbPrt = "" & Trim(!PARTNUM)
            lblDsc = "" & Trim(!PADESC)
            lblTyp = Format(0 + !PALEVEL, "#")
         End With
         GetPart = True
      Else
         MsgBox "Part Wasn't Found.", vbExclamation, Caption
         cmbPrt = ""
         lblDsc = ""
         GetPart = False
      End If
      On Error Resume Next
      rdoPrt.Close
   Else
      sPartNumber = ""
      cmbPrt = ""
   End If
   Exit Function
   
   DiaErr1:
   sProcName = "getpart"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function



Private Sub optDis_Click()
   PrintReport
   
End Sub

Private Sub optPrn_Click()
   PrintReport
   
End Sub


Private Sub txtBeg_DropDown()
   ShowCalendar Me
   
End Sub

Private Sub txtBeg_LostFocus()
   If Len(Trim(txtBeg)) = 0 Then txtBeg = "ALL"
   If txtBeg <> "ALL" Then txtBeg = CheckDate(txtBeg)
   
End Sub


Private Sub txtEnd_DropDown()
   ShowCalendar Me
   
End Sub

Private Sub txtEnd_LostFocus()
   If Len(txtEnd) > 3 Then
      txtEnd = CheckDate(txtEnd)
   Else
      txtEnd = "ALL"
   End If
   
End Sub
