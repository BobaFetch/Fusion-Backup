VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaWip
   BorderStyle = 3 'Fixed Dialog
   Caption = "Work In Process (Report)"
   ClientHeight = 3615
   ClientLeft = 2115
   ClientTop = 1125
   ClientWidth = 6810
   ClipControls = 0 'False
   ControlBox = 0 'False
   ForeColor = &H00000000&
   LinkTopic = "Form1"
   MDIChild = -1 'True
   MinButton = 0 'False
   PaletteMode = 1 'UseZOrder
   ScaleHeight = 3615
   ScaleWidth = 6810
   ShowInTaskbar = 0 'False
   Begin VB.ComboBox cmbcls
      Height = 315
      Left = 1320
      TabIndex = 15
      Tag = "8"
      Top = 1200
      Width = 1095
   End
   Begin VB.ComboBox txtAso
      Height = 315
      Left = 1320
      TabIndex = 0
      Tag = "4"
      Top = 720
      Width = 1095
   End
   Begin VB.CheckBox optDsc
      Caption = "____"
      ForeColor = &H8000000F&
      Height = 195
      Left = 3000
      TabIndex = 1
      Top = 2040
      Width = 855
   End
   Begin VB.CheckBox optExt
      Caption = "____"
      ForeColor = &H8000000F&
      Height = 195
      Left = 3000
      TabIndex = 2
      Top = 2280
      Width = 975
   End
   Begin VB.CommandButton cmdCan
      Cancel = -1 'True
      Caption = "Close"
      Height = 360
      Left = 5640
      TabIndex = 7
      TabStop = 0 'False
      Top = 0
      Width = 1065
   End
   Begin VB.Frame fraPrn
      BorderStyle = 0 'None
      Height = 495
      Left = 5640
      TabIndex = 4
      Top = 360
      Width = 1095
      Begin VB.CommandButton optDis
         Height = 330
         Left = 0
         Picture = "diaPin06.frx":0000
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
         Picture = "diaPin06.frx":017E
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
      TabIndex = 3
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
      PictureUp = "diaPin06.frx":0308
      PictureDn = "diaPin06.frx":044E
   End
   Begin ResizeLibCtl.ReSize ReSize1
      Left = 4920
      Top = 0
      _Version = 196615
      _ExtentX = 741
      _ExtentY = 741
      _StockProps = 0
      Enabled = -1 'True
      FormMinWidth = 0
      FormMinHeight = 0
      FormDesignHeight = 3615
      FormDesignWidth = 6810
   End
   Begin Threed.SSRibbon ShowPrinters
      Height = 255
      Left = 360
      TabIndex = 11
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
      PictureUp = "diaPin06.frx":0594
      PictureDn = "diaPin06.frx":06DA
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "As Of "
      Height = 285
      Index = 0
      Left = 240
      TabIndex = 16
      Top = 720
      Width = 1065
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "(Blank For ALL) "
      Height = 285
      Index = 10
      Left = 2760
      TabIndex = 14
      Top = 1200
      Width = 2385
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Part Class"
      Height = 285
      Index = 2
      Left = 240
      TabIndex = 13
      Top = 1200
      Width = 1065
   End
   Begin VB.Label lblPrinter
      Appearance = 0 'Flat
      BorderStyle = 1 'Fixed Single
      Caption = "Default Printer"
      ForeColor = &H00800000&
      Height = 255
      Left = 720
      TabIndex = 12
      Top = 0
      Width = 2760
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Include:"
      Height = 285
      Index = 5
      Left = 240
      TabIndex = 10
      Top = 1800
      Width = 1785
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Extended Descriptions?"
      Height = 285
      Index = 3
      Left = 360
      TabIndex = 9
      Top = 2280
      Width = 1815
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Part Descriptions?"
      Height = 285
      Index = 6
      Left = 360
      TabIndex = 8
      Top = 2040
      Width = 1785
   End
End
Attribute VB_Name = "diaWip"
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

'*************************************************************************************
' diaWip - Work In Progress (Report)
'
' Notes:
'
' Created: 11/05/02 (nth)
' Revisions:
'   11/13/02 (nth) Modified form to support all three reports.
'   12/03/02 (nth) Added partclass combo.
'
'*************************************************************************************

Dim bOnLoad As Byte
Dim bCancel As Byte
Public bCosting As Byte 'Actual=0 / Standard=1

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'*************************************************************************************

Private Sub cmbCls_LostFocus()
   If Trim(cmbcls) = "" Then cmbcls = "ALL"
End Sub

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, _
                             X As Single, Y As Single)
   bCancel = True
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      SelectHelpTopic Me, "Reports"
      MouseCursor 0
      cmdHlp = False
   End If
End Sub

Private Sub Form_Activate()
   On Error Resume Next
   If bOnLoad Then
      FillCombo
      bOnLoad = False
   End If
   MdiSect.lblBotPanel = Caption
   MouseCursor 0
End Sub

Private Sub Form_Load()
   SetDiaPos Me
   FormatControls
   sCurrForm = Caption
   GetOptions
   bOnLoad = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveOptions
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set diaWip = Nothing
End Sub

Private Sub PrintReport()
   Dim sWindows As String
   MouseCursor 13
   
   On Error GoTo DiaErr1
   ReopenJet
   
   sWindows = GetWindowsDir()
   SetMdiReportsize MdiSect
   
   MdiSect.crw.ReportFileName = sReportPath & "finwip.rpt"
   
   MdiSect.crw.Formulas(0) = "CompanyName='" & sFacility & "'"
   MdiSect.crw.Formulas(1) = "RequestBy='Requested By: " & sInitials & "'"
   MdiSect.crw.Formulas(2) = "Title1='Work In Process Inventory As Of " _
                        & Format(txtAso, "mm/dd/yy") & "'"
   
   If optDsc.Value = vbUnchecked Then
      MdiSect.crw.Formulas(3) = "ShowDsc='0'"
   Else
      MdiSect.crw.Formulas(3) = "ShowDsc='1'"
   End If
   
   If optExt.Value = vbUnchecked Then
      MdiSect.crw.Formulas(4) = "ShowExt='0'"
   Else
      MdiSect.crw.Formulas(5) = "ShowExt='1'"
   End If
   
   SetCrystalAction Me
   
   optPrn.Enabled = True
   optDis.Enabled = True
   MouseCursor 0
   Exit Sub
   
   DiaErr1:
   sProcName = "printreport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Public Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Public Sub SaveOptions()
   Dim sOptions As String
   SaveSetting "Esi2000", "EsiFina", "finwip", sOptions
End Sub

Public Sub GetOptions()
   Dim sOptions As String
   ' Get By Menu Option
   sOptions = GetSetting("Esi2000", "EsiFina", "finwip", sOptions)
End Sub

Private Sub optDis_Click()
   PrintReport
End Sub

Private Sub optDsc_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
End Sub

Private Sub optPrn_Click()
   PrintReport
End Sub

Private Sub ShowPrinters_Click(Value As Integer)
   SysPrinters.Show
   ShowPrinters.Value = False
End Sub

Public Sub FillCombo()
   FillProductClasses Me
   If Trim(cmbcls) = "" Then cmbcls = "ALL"
End Sub

Private Sub txtAso_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtAso_LostFocus()
   txtAso = CheckDate(txtAso)
End Sub
