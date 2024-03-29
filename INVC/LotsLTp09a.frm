VERSION 5.00
Begin VB.Form LotsLTp09a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ITAR and EAR Report"
   ClientHeight    =   4275
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   8115
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4275
   ScaleWidth      =   8115
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox optEAR 
      Height          =   195
      Left            =   2400
      TabIndex        =   26
      Top             =   2320
      Width           =   240
   End
   Begin VB.CheckBox optITAR 
      Height          =   195
      Left            =   2400
      TabIndex        =   24
      Top             =   2040
      Width           =   240
   End
   Begin VB.ComboBox cboClass 
      DataSource      =   "rDt1"
      Height          =   315
      Left            =   2400
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Tag             =   "9"
      ToolTipText     =   "Contains Part Numbers With Lots"
      Top             =   1560
      Width           =   855
   End
   Begin VB.CheckBox chkShowObsoleteParts 
      Height          =   195
      Left            =   2400
      TabIndex        =   9
      Top             =   3840
      Width           =   240
   End
   Begin VB.CheckBox chkShowInactiveParts 
      Height          =   195
      Left            =   2400
      TabIndex        =   8
      Top             =   3540
      Width           =   240
   End
   Begin VB.CheckBox chkShowLotDetails 
      Height          =   195
      Left            =   2400
      TabIndex        =   7
      Top             =   3240
      Width           =   240
   End
   Begin VB.CheckBox chkShowExtDesc 
      Height          =   195
      Left            =   2400
      TabIndex        =   6
      Top             =   2940
      Width           =   240
   End
   Begin VB.CheckBox chkShowDesc 
      Height          =   195
      Left            =   2400
      TabIndex        =   5
      Top             =   2640
      Width           =   240
   End
   Begin VB.ComboBox cboPart 
      DataSource      =   "rDt1"
      Height          =   315
      Left            =   2400
      TabIndex        =   0
      Tag             =   "99"
      ToolTipText     =   "Contains Part Numbers With Lots"
      Top             =   840
      Width           =   3135
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "LotsLTp09a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   16
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox cboThroughDate 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   4440
      TabIndex        =   3
      Tag             =   "4"
      Top             =   1200
      Width           =   1095
   End
   Begin VB.ComboBox cboFromDate 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   2400
      TabIndex        =   1
      Tag             =   "4"
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6840
      TabIndex        =   10
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6840
      TabIndex        =   2
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "LotsLTp09a.frx":07AE
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   600
         Picture         =   "LotsLTp09a.frx":092C
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.Label z 
      BackStyle       =   0  'Transparent
      Caption         =   "Include Only EAR"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   27
      Top             =   2320
      Width           =   2085
   End
   Begin VB.Label z 
      BackStyle       =   0  'Transparent
      Caption         =   "Include ITAR"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   25
      Top             =   2040
      Width           =   2085
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Actual Transaction Date)"
      Height          =   285
      Index           =   1
      Left            =   6000
      TabIndex        =   23
      Top             =   1260
      Width           =   1935
   End
   Begin VB.Label z 
      BackStyle       =   0  'Transparent
      Caption         =   "Show Obsolete Parts"
      Height          =   285
      Index           =   3
      Left            =   240
      TabIndex        =   22
      Top             =   3840
      Width           =   2085
   End
   Begin VB.Label z 
      BackStyle       =   0  'Transparent
      Caption         =   "Show Inactive Parts"
      Height          =   285
      Index           =   6
      Left            =   240
      TabIndex        =   21
      Top             =   3540
      Width           =   2085
   End
   Begin VB.Label z 
      BackStyle       =   0  'Transparent
      Caption         =   "Show Lot Details"
      Height          =   285
      Index           =   5
      Left            =   240
      TabIndex        =   20
      Top             =   3240
      Width           =   1785
   End
   Begin VB.Label z 
      BackStyle       =   0  'Transparent
      Caption         =   "Show Extended Description"
      Height          =   285
      Index           =   4
      Left            =   240
      TabIndex        =   19
      Top             =   2940
      Width           =   2085
   End
   Begin VB.Label z 
      BackStyle       =   0  'Transparent
      Caption         =   "Show Description"
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   18
      Top             =   2640
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Class(es)"
      Height          =   285
      Index           =   5
      Left            =   240
      TabIndex        =   17
      Top             =   1620
      Width           =   1815
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Through"
      Height          =   255
      Index           =   4
      Left            =   3600
      TabIndex        =   15
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include lots from"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   14
      Top             =   1260
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number(s)"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   13
      Top             =   900
      Width           =   1815
   End
End
Attribute VB_Name = "LotsLTp09a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Stanwood, Washington, USA  ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'3/1/05 Changed date handling
'5/16/05 corrected group show/hide
'9/15/05 Added Inventory Transfer to report table (32)
Option Explicit
Dim bOnLoad As Byte

Dim iProg As Integer
Dim iTotalLots As Integer
Dim sLots(1000) As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cboClass_KeyPress(KeyAscii As Integer)
   
   'if backspace or space, go to first entry (ALL)
   If KeyAscii = 8 Or KeyAscii = 32 Then
      cboClass.ListIndex = 0
   End If
   
End Sub

Private Sub cboPart_DropDown()
   
   ' if part exists in list, don't repopulate
   If cboPart.ListIndex <> -1 And cboPart.Text <> "<ALL>" Then
      Exit Sub
   End If
   Dim part As New ClassPart
   part.PopulatePartComboTest cboPart, True
End Sub

Private Sub cboPart_KeyPress(KeyAscii As Integer)
   'if backspace or space, go to first entry (ALL)
   If KeyAscii = 8 Or KeyAscii = 32 Then
      cboPart.ListIndex = 0
      cboPart.Text = cboPart.List(0)
   End If
End Sub

Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 907
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub Form_Activate()
   On Error Resume Next
   MdiSect.lblBotPanel = Caption
   MouseCursor 0
   If bOnLoad Then
      cboPart.SetFocus
      bOnLoad = 0
      Dim cls As New ClassPartClass
      cls.PopulatePartClassCombo cboClass, True
      cboPart.Clear
      cboPart.AddItem "<ALL>"
      cboPart.ListIndex = 0
   End If
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   GetOptions
   bOnLoad = 1
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveOptions
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set LotsLTp01a = Nothing
   
End Sub

Private Sub PrintReport()
   Dim sBook As String
   MouseCursor 13
   
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   Dim sqlITAR As String
   Dim sqlEAR As String
   Dim sSql As String
   
   On Error GoTo DiaErr1
   
   Dim part As String, cls As String
   part = Trim(cboPart)
   If part = "<ALL>" Then
      part = "ALL"
   End If
   cls = Trim(cboClass)
   If cls = "<ALL>" Then
      cls = "ALL"
   End If
   
   
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   sCustomReport = GetCustomReport("invlt09")
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "Includes"
   aFormulaName.Add "RequestBy"
   aFormulaName.Add "PartNumber"
   aFormulaName.Add "FromDate"
   
   aFormulaName.Add "ThroughDate"
   aFormulaName.Add "PartClass"
   aFormulaName.Add "ShowDescription"
   aFormulaName.Add "ShowExtDescription"
   aFormulaName.Add "ShowLotDetails"
   aFormulaName.Add "ShowInactiveParts"
   aFormulaName.Add "ShowObsoleteParts"
   aFormulaName.Add "IncOnlyITAR"
   aFormulaName.Add "IncOnlyEAR"
   
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'" & part _
                         & "... From " & cboFromDate & " Through  " & cboThroughDate & "'")
   aFormulaValue.Add CStr("'Requested By:" & CStr(sInitials) & "'")
   
   aFormulaValue.Add CStr("'" & Compress(part) & "'")
   aFormulaValue.Add CStr("'" & cboFromDate.Text & "'")
   aFormulaValue.Add CStr("'" & cboThroughDate.Text & "'")
  
   If cboClass = "" Then
     aFormulaValue.Add CStr("'ALL'")
   Else
     aFormulaValue.Add CStr("'" & cls & "'")
   End If
  
   aFormulaValue.Add chkShowDesc.Value
   aFormulaValue.Add chkShowExtDesc.Value
   aFormulaValue.Add chkShowLotDetails.Value
   aFormulaValue.Add chkShowInactiveParts.Value
   aFormulaValue.Add chkShowObsoleteParts.Value
  
   aFormulaValue.Add optITAR.Value
   aFormulaValue.Add optEAR.Value
  
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   
   sSql = cCRViewer.GetReportSelectionFormula
   
   If (sSql <> "") Then
      If (optITAR.Value = 1) Then
         sqlITAR = "{PartTable.PAITARRPT} = 1"
      End If
      
      If (optEAR.Value = 1) Then
         sqlEAR = "{PartTable.PAEARRPT} = 1"
      End If

      If (sqlITAR <> "") Then
         sSql = sSql & " AND (" & sqlITAR
      End If
      
      If (sqlEAR <> "") Then
         If (sqlITAR <> "") Then
            sSql = sSql & " OR " & sqlEAR & ")"
         Else
            sSql = sSql & " AND (" & sqlEAR & ")"
         End If
      Else
         If (sqlITAR <> "") Then
            sSql = sSql & ")"
         End If
      End If
   
      cCRViewer.SetReportSelectionFormula (sSql)
   
   End If
   
   cCRViewer.CRViewerSize Me
   cCRViewer.SetDbTableConnection
   cCRViewer.ShowGroupTree False
   cCRViewer.OpenCrystalReportObject Me, aFormulaName
 
   cCRViewer.ClearFieldCollection aFormulaName
   cCRViewer.ClearFieldCollection aFormulaValue
  
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "printreport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   cboThroughDate = Format(ES_SYSDATE, "mm/dd/yy")
   cboFromDate = Left(cboThroughDate, 3) & "01" & Right(cboThroughDate, 3)
End Sub

Private Sub SaveOptions()
   Dim sOptions As String
   sOptions = CStr(chkShowDesc.Value) + CStr(chkShowExtDesc.Value) + CStr(chkShowLotDetails.Value) _
      + CStr(chkShowInactiveParts.Value) + CStr(chkShowObsoleteParts.Value) + "000000000"
   SaveSetting "Esi2000", "EsiInvc", "lt07", sOptions
   
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiInvc", "lt07", "0000000000000000")
   If Len(sOptions) > 7 Then
      chkShowDesc.Value = Val(Mid(sOptions, 3, 1))
      chkShowExtDesc.Value = Val(Mid(sOptions, 4, 1))
      chkShowLotDetails.Value = Val(Mid(sOptions, 5, 1))
      chkShowInactiveParts.Value = Val(Mid(sOptions, 6, 1))
      chkShowObsoleteParts.Value = Val(Mid(sOptions, 7, 1))
   End If
End Sub

Private Sub optDet_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optDis_Click()
   PrintReport
   
End Sub


Private Sub optPrn_Click()
   PrintReport
   
End Sub




Private Sub cboFromDate_DropDown()
   ShowCalendar Me
   
End Sub


Private Sub cboFromDate_LostFocus()
   If Len(Trim(cboFromDate)) = 0 Then cboFromDate = "ALL"
   If cboFromDate <> "ALL" Then cboFromDate = CheckDate(cboFromDate)
   
End Sub


Private Sub cboThroughDate_DropDown()
   ShowCalendar Me
End Sub

Private Sub cboThroughDate_LostFocus()
   If Len(Trim(cboThroughDate)) = 0 Then cboThroughDate = "ALL"
   If cboThroughDate <> "ALL" Then cboThroughDate = CheckDate(cboThroughDate)
End Sub


Private Sub cboPart_LostFocus()
   cboPart = CheckLen(cboPart, 30)
   If Trim(cboPart) = "" Then cboPart = "<ALL>"
   
   'if an individual part is selected, disable the class combo box
   If cboPart <> "<ALL>" Then
      If cboClass.ListCount > 0 Then
         cboClass.ListIndex = 0
      End If
      cboClass.Enabled = False
   Else
      cboClass.Enabled = True
   End If
End Sub

'Private Sub txtClass_LostFocus()
'   If Trim(txtClass.Text) = "" Then
'      txtClass.Text = "ALL"
'   End If
'End Sub
