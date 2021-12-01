VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaGARMAN
   BorderStyle = 3 'Fixed Dialog
   Caption = "Part Status (Report)"
   ClientHeight = 3030
   ClientLeft = 2115
   ClientTop = 1125
   ClientWidth = 7170
   ClipControls = 0 'False
   ControlBox = 0 'False
   ForeColor = &H80000007&
   LinkTopic = "Form1"
   MDIChild = -1 'True
   MinButton = 0 'False
   PaletteMode = 1 'UseZOrder
   ScaleHeight = 3030
   ScaleWidth = 7170
   ShowInTaskbar = 0 'False
   Begin VB.CheckBox optVew
      Height = 255
      Left = 3600
      TabIndex = 18
      Top = 0
      Visible = 0 'False
      Width = 735
   End
   Begin VB.CheckBox optCom
      Caption = "____"
      ForeColor = &H8000000F&
      Height = 255
      Left = 2040
      TabIndex = 3
      Top = 2160
      Width = 735
   End
   Begin VB.ComboBox txtEnd
      Height = 315
      Left = 3960
      TabIndex = 2
      Tag = "4"
      Top = 1440
      Width = 1095
   End
   Begin VB.ComboBox txtBeg
      Height = 315
      Left = 2040
      TabIndex = 1
      Tag = "4"
      Top = 1440
      Width = 1095
   End
   Begin VB.CommandButton cmdFnd
      Height = 315
      Left = 5160
      Picture = "diaGARMAN.frx":0000
      Style = 1 'Graphical
      TabIndex = 14
      TabStop = 0 'False
      ToolTipText = "Find A Part"
      Top = 600
      UseMaskColor = -1 'True
      Width = 350
   End
   Begin VB.TextBox cmbPrt
      Height = 285
      Left = 2040
      TabIndex = 0
      Tag = "3"
      Top = 600
      Width = 3015
   End
   Begin VB.CommandButton cmdCan
      Cancel = -1 'True
      Caption = "Close"
      Height = 360
      Left = 6000
      TabIndex = 7
      TabStop = 0 'False
      Top = 0
      Width = 1065
   End
   Begin VB.Frame fraPrn
      BorderStyle = 0 'None
      Height = 615
      Left = 6000
      TabIndex = 4
      Top = 360
      Width = 1215
      Begin VB.CommandButton optDis
         Height = 330
         Left = 0
         Picture = "diaGARMAN.frx":0342
         Style = 1 'Graphical
         TabIndex = 5
         ToolTipText = "Display The Report"
         Top = 120
         UseMaskColor = -1 'True
         Width = 495
      End
      Begin VB.CommandButton optPrn
         Height = 330
         Left = 560
         Picture = "diaGARMAN.frx":04C0
         Style = 1 'Graphical
         TabIndex = 6
         ToolTipText = "Print The Report"
         Top = 120
         UseMaskColor = -1 'True
         Width = 495
      End
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
      PictureUp = "diaGARMAN.frx":064A
      PictureDn = "diaGARMAN.frx":0790
   End
   Begin Threed.SSRibbon ShowPrinters
      Height = 255
      Left = 360
      TabIndex = 10
      ToolTipText = "Show System Printers"
      Top = 0
      Width = 255
      _Version = 65536
      _ExtentX = 450
      _ExtentY = 450
      _StockProps = 65
      BackColor = 12632256
      GroupAllowAllUp = -1 'True
      RoundedCorners = 0 'False
      BevelWidth = 0
      Outline = 0 'False
      PictureUp = "diaGARMAN.frx":08D6
      PictureDn = "diaGARMAN.frx":0A1C
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Include:"
      Height = 285
      Index = 4
      Left = 120
      TabIndex = 17
      Top = 1920
      Width = 1665
   End
   Begin VB.Label lblDsc
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      Height = 285
      Left = 2040
      TabIndex = 16
      Top = 960
      Width = 3015
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Status Comments"
      Height = 285
      Index = 3
      Left = 120
      TabIndex = 15
      Top = 2160
      Width = 1665
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Through"
      Height = 285
      Index = 2
      Left = 3240
      TabIndex = 13
      Top = 1440
      Width = 1785
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "With Dates From "
      Height = 285
      Index = 1
      Left = 120
      TabIndex = 12
      Top = 1440
      Width = 1785
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Status Of Part Number"
      Height = 285
      Index = 0
      Left = 120
      TabIndex = 11
      Top = 600
      Width = 1785
   End
   Begin VB.Label lblPrinter
      Appearance = 0 'Flat
      BorderStyle = 1 'Fixed Single
      Caption = "Default Printer"
      ForeColor = &H00800000&
      Height = 255
      Left = 720
      TabIndex = 8
      Top = 0
      Width = 2760
   End
End
Attribute VB_Name = "diaGARMAN"
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

'*********************************************************************************
' diaGARMAN
'
' Notes: Created for GARMAN
'
' Created: 06/29/04 (nth)
' Revisions:
'
'*********************************************************************************

Dim bOnload As Byte
Dim bCancel As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'*********************************************************************************


Private Sub cmbPrt_GotFocus()
   SelectFormat Me
End Sub


Private Sub cmbPrt_LostFocus()
   cmbPrt = CheckLen(cmbPrt, 30)
   FindPart Me, cmbPrt
End Sub

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, _
                             X As Single, Y As Single)
   bCancel = True
End Sub

Private Sub cmdFnd_Click()
   optVew.Value = vbChecked
   VewParts.Show
End Sub

Private Sub Form_Activate()
   On Error Resume Next
   MdiSect.lblBotPanel = Caption
   If bOnload Then
      FillCombo
      bOnload = False
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   SetDiaPos Me
   FormatControls
   sCurrForm = Caption
   txtBeg = Format(ES_SYSDATE, "mm/01/yy")
   txtEnd = Format(ES_SYSDATE, "mm/dd/yy")
   GetOptions
   bOnload = True
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set diaGARMAN = Nothing
End Sub

Private Sub optDis_Click()
   PrintReport
End Sub

Private Sub optPrn_Click()
   PrintReport
End Sub

Private Sub ShowPrinters_Click(Value As Integer)
   SysPrinters.Show
   ShowPrinters.Value = False
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Private Sub PrintReport()
   Dim sCustomReport As String
   Dim sPart As String
   MouseCursor 13
   
   On Error GoTo DiaErr1
   
   optPrn.Enabled = False
   optDis.Enabled = False
   
   sPart = Compress(cmbPrt)
   
   SetMdiReportsize MdiSect
   
   sCustomReport = GetCustomReport("garman.rpt")
   MdiSect.crw.ReportFileName = sReportPath & sCustomReport
   
   MdiSect.crw.Formulas(0) = "CompanyName='" & sFacility & "'"
   MdiSect.crw.Formulas(1) = "Status=" & Val(optCom)
   MdiSect.crw.Formulas(2) = "RequestBy='Requested By: " _
                        & Secure.UserInitials & "'"
   
   sSql = "({SoitTable.ITSCHED} >= #" & txtBeg & "# AND " _
          & "{SoitTable.ITSCHED} <= #" & txtEnd & "# AND " _
          & "ISNULL({SoitTable.ITACTUAL}) AND " _
          & "{PartTable.PARTREF} = '" & sPart & "') OR " _
          & "({SoitTable.ITSCHED} >= #" & txtBeg & "# AND " _
          & "{SoitTable.ITSCHED} <= #" & txtEnd & "# AND " _
          & "{SoitTable.ITACTUAL} > #" & txtEnd & "# AND " _
          & "{PartTable.PARTREF} = '" & sPart & "')"
   MdiSect.crw.SelectionFormula = sSql
   
   SetCrystalAction Me
   optPrn.Enabled = True
   optDis.Enabled = True
   MouseCursor 0
   Exit Sub
   
   DiaErr1:
   optPrn.Enabled = False
   optDis.Enabled = False
   sProcName = "printrep"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub FillCombo()
   On Error GoTo DiaErr1
   Exit Sub
   DiaErr1:
   sProcName = "fillcomb"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub SaveOptions()
   Dim sOptions As String
End Sub

Private Sub GetOptions()
   Dim sOptions As String
End Sub

Private Sub txtBeg_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtBeg_GotFocus()
   SelectFormat Me
End Sub

Private Sub txtBeg_LostFocus()
   txtBeg = CheckDate(txtBeg)
End Sub

Private Sub txtEnd_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtEnd_GotFocus()
   SelectFormat Me
End Sub

Private Sub txtEnd_LostFocus()
   txtEnd = CheckDate(txtEnd)
End Sub
