VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form diaRMFG 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Raw Material/Finished Goods Inventory"
   ClientHeight    =   3675
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   6765
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H80000007&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3675
   ScaleWidth      =   6765
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox optQOH 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2160
      TabIndex        =   24
      Top             =   2760
      Width           =   855
   End
   Begin VB.CheckBox optLot 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2160
      TabIndex        =   8
      Top             =   3360
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ComboBox cmbCls 
      Height          =   315
      Left            =   2160
      TabIndex        =   5
      Tag             =   "8"
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CheckBox optDsc 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2160
      TabIndex        =   6
      Top             =   2280
      Width           =   855
   End
   Begin VB.CheckBox optExt 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   285
      Left            =   2160
      TabIndex        =   7
      Top             =   2520
      Width           =   855
   End
   Begin VB.ComboBox txtDte 
      Height          =   315
      Left            =   2160
      TabIndex        =   0
      Tag             =   "4"
      Top             =   600
      Width           =   1095
   End
   Begin VB.CheckBox typ 
      Caption         =   "4"
      Height          =   255
      Index           =   4
      Left            =   3600
      TabIndex        =   4
      Top             =   1080
      Value           =   1  'Checked
      Width           =   495
   End
   Begin VB.CheckBox typ 
      Caption         =   "3"
      Height          =   255
      Index           =   3
      Left            =   3120
      TabIndex        =   3
      Top             =   1080
      Value           =   1  'Checked
      Width           =   495
   End
   Begin VB.CheckBox typ 
      Caption         =   "2"
      Height          =   255
      Index           =   2
      Left            =   2640
      TabIndex        =   2
      Top             =   1080
      Value           =   1  'Checked
      Width           =   495
   End
   Begin VB.CheckBox typ 
      Caption         =   "1"
      Height          =   255
      Index           =   1
      Left            =   2160
      TabIndex        =   1
      Top             =   1080
      Value           =   1  'Checked
      Width           =   495
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   5640
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   5640
      TabIndex        =   11
      Top             =   360
      Width           =   1215
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "diaRMFG.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "diaRMFG.frx":017E
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   14
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
      PictureUp       =   "diaRMFG.frx":0308
      PictureDn       =   "diaRMFG.frx":044E
   End
   Begin Threed.SSRibbon ShowPrinters 
      Height          =   255
      Left            =   360
      TabIndex        =   15
      ToolTipText     =   "Show System Printers"
      Top             =   0
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   450
      _StockProps     =   65
      BackColor       =   12632256
      GroupAllowAllUp =   -1  'True
      RoundedCorners  =   0   'False
      BevelWidth      =   0
      Outline         =   0   'False
      PictureUp       =   "diaRMFG.frx":0594
      PictureDn       =   "diaRMFG.frx":06DA
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "No QOH"
      Height          =   285
      Index           =   7
      Left            =   120
      TabIndex        =   25
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Lots On Hand"
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   23
      Top             =   3360
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Class"
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   22
      Top             =   1560
      Width           =   1065
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For ALL) "
      Height          =   285
      Index           =   10
      Left            =   3600
      TabIndex        =   21
      Top             =   1560
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   285
      Index           =   5
      Left            =   120
      TabIndex        =   20
      Top             =   2040
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Extended Descriptions"
      Height          =   285
      Index           =   3
      Left            =   120
      TabIndex        =   19
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Descriptions"
      Height          =   285
      Index           =   6
      Left            =   120
      TabIndex        =   18
      Top             =   2280
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "As Of Date"
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   17
      Top             =   600
      Width           =   1185
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Types"
      Height          =   285
      Index           =   4
      Left            =   120
      TabIndex        =   16
      Top             =   1080
      Width           =   1305
   End
   Begin VB.Label lblPrinter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Default Printer"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   720
      TabIndex        =   13
      Top             =   0
      Width           =   2760
   End
End
Attribute VB_Name = "diaRMFG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of                     ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***

'See the UpdateTables prodecure for database revisions

'*********************************************************************************
' diaRMFG - Raw Material Finished Goods
'
' Notes:
'
' Created: 02/09/04 (nth)
' Revisions:
'   02/12/04 (nth) Added part class and lot detail options.
'   05/04/04 (nth) Added include parts with no QOH option.
'
'*********************************************************************************

Option Explicit

Dim bOnLoad As Byte
Dim bCancel As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'*********************************************************************************

Private Sub cmbCls_LostFocus()
   If Trim(cmbCls) = "" Then cmbCls = "ALL"
End Sub

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, _
                             X As Single, Y As Single)
   bCancel = True
End Sub

Private Sub Form_Activate()
   On Error Resume Next
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillCombo
      bOnLoad = False
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   FormLoad Me
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
   Set diaRMFG = Nothing
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
   Dim sCustomerReport As String
   Dim sType As String
   Dim b As Byte
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   Dim cCRViewer As EsCrystalRptViewer
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport

   
   MouseCursor 13
   On Error GoTo DiaErr1
   
   If Trim(cmbCls) = "" Then cmbCls = "ALL"
   
   'SetMdiReportsize MdiSect
   
   For b = 1 To 4
      If typ(b) = vbChecked Then
         sType = sType & CStr(b) & ","
      End If
   Next
   If Len(sType) Then
      sType = Left(sType, Len(sType) - 1)
   End If
   
   If optLot = vbChecked Then
      sCustomReport = GetCustomReport("finRMFGa.rpt")
   Else
      sCustomReport = GetCustomReport("finRMFGb.rpt")
   End If
'   MdiSect.crw.ReportFileName = sReportPath & sCustomReport
'
'   MdiSect.crw.Formulas(0) = "CompanyName='" & sFacility & "'"
'   MdiSect.crw.Formulas(1) = "RequestBy='Requested By: " _
'                        & sInitials & "'"
'   MdiSect.crw.Formulas(2) = "AsOf='" & txtDte & "'"
'   MdiSect.crw.Formulas(3) = "Title1='As Of " & txtDte & "'"
'   MdiSect.crw.Formulas(4) = "Title2='Includes Part Types " & sType & "'"
'   MdiSect.crw.Formulas(5) = "Title3='For Part Class " & cmbCls & "'"
'
'   MdiSect.crw.Formulas(6) = "Dsc=" & optDsc
'   MdiSect.crw.Formulas(7) = "Ext=" & optExt
'   MdiSect.crw.Formulas(8) = "QOH=" & optQOH
   
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "RequestBy"
   aFormulaName.Add "AsOf"
   aFormulaName.Add "Title1"
   aFormulaName.Add "Title2"
   aFormulaName.Add "Title3"
   aFormulaName.Add "Dsc"
   aFormulaName.Add "Ext"
   aFormulaName.Add "QOH="

   aFormulaValue.Add sFacility
   aFormulaValue.Add sInitials
   aFormulaValue.Add txtDte
   aFormulaValue.Add txtDte
   aFormulaValue.Add "Includes Part Types " & sType
   aFormulaValue.Add "For Part Class " & cmbCls
   aFormulaValue.Add optDsc
   aFormulaValue.Add optExt
   aFormulaValue.Add optQOH
   
   sSql = "{InvaTable.INADATE}<=cdate('" & txtDte & _
          "') AND {PartTable.PALEVEL} IN [" & sType & "]"
   If UCase(cmbCls) <> "ALL" Then
      sSql = sSql & " AND {PartTable.PACLASS}='" & cmbCls & "'"
   End If
   'If optLot = vbUnchecked Then
   '    sSql = sSql & " AND {PartTable.PAQOH} > 0"
   'End If
   'MdiSect.crw.SelectionFormula = sSql
   
   'SetCrystalAction Me
   
    cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
    cCRViewer.CRViewerSize Me
    ' Set report parameter
    cCRViewer.SetDbTableConnection
    ' print the copies
    cCRViewer.SetReportSelectionFormula sSql
    cCRViewer.OpenCrystalReportObject Me, aFormulaName
    cCRViewer.ShowGroupTree False
   
    
    cCRViewer.ClearFieldCollection aFormulaName
    cCRViewer.ClearFieldCollection aFormulaValue
   
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "printrep"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub FillCombo()
   FillProductClasses Me
End Sub

Private Sub txtDte_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtDte_LostFocus()
   txtDte = CheckDate(txtDte)
End Sub

Private Sub SaveOptions()
   Dim sOptions As String
   
   ' Save by Menu Option
   sOptions = RTrim(typ(1).Value) _
              & RTrim(typ(2).Value) _
              & RTrim(typ(3).Value) _
              & RTrim(typ(4).Value) _
              & RTrim(optDsc.Value) _
              & RTrim(optExt.Value) _
              & RTrim(optLot.Value) _
              & RTrim(optQOH.Value)
   SaveSetting "Esi2000", "EsiFina", Me.Name, Trim(sOptions)
   SaveSetting "Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiFina", Me.Name, sOptions)
   If Len(Trim(sOptions)) > 0 Then
      typ(1).Value = Val(Mid(sOptions, 1, 1))
      typ(2).Value = Val(Mid(sOptions, 2, 1))
      typ(3).Value = Val(Mid(sOptions, 3, 1))
      typ(4).Value = Val(Mid(sOptions, 4, 1))
      optDsc.Value = Val(Mid(sOptions, 5, 1))
      optExt.Value = Val(Mid(sOptions, 6, 1))
      optLot.Value = Val(Mid(sOptions, 7, 1))
      optQOH.Value = Val(Mid(sOptions, 8, 1))
   Else
      typ(1).Value = vbChecked
      typ(2).Value = vbChecked
      typ(3).Value = vbChecked
      typ(4).Value = vbChecked
      optDsc.Value = vbUnchecked
      optExt.Value = vbUnchecked
      optLot.Value = vbUnchecked
      optQOH.Value = vbUnchecked
   End If
   lblPrinter = GetSetting("Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter)
   If lblPrinter = "" Then
      lblPrinter = "Default Printer"
   End If
End Sub
