VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form PurcPRp06a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Purchase Expediting Report"
   ClientHeight    =   4020
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7200
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4020
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbGroupBy 
      Height          =   315
      Index           =   1
      ItemData        =   "PurcPRp06a.frx":0000
      Left            =   2160
      List            =   "PurcPRp06a.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   30
      Top             =   3480
      Width           =   1935
   End
   Begin VB.CheckBox chkType 
      Caption         =   "8"
      Height          =   255
      Index           =   8
      Left            =   6360
      TabIndex        =   9
      Top             =   2280
      Width           =   435
   End
   Begin VB.CheckBox chkType 
      Caption         =   "7"
      Height          =   255
      Index           =   7
      Left            =   5760
      TabIndex        =   8
      Top             =   2280
      Width           =   435
   End
   Begin VB.CheckBox chkType 
      Caption         =   "6"
      Height          =   255
      Index           =   6
      Left            =   5160
      TabIndex        =   7
      Top             =   2280
      Width           =   435
   End
   Begin VB.CheckBox chkType 
      Caption         =   "5"
      Height          =   255
      Index           =   5
      Left            =   4560
      TabIndex        =   6
      Top             =   2280
      Width           =   435
   End
   Begin VB.CheckBox chkType 
      Caption         =   "4"
      Height          =   255
      Index           =   4
      Left            =   3960
      TabIndex        =   5
      Top             =   2280
      Width           =   435
   End
   Begin VB.CheckBox chkType 
      Caption         =   "3"
      Height          =   255
      Index           =   3
      Left            =   3360
      TabIndex        =   4
      Top             =   2280
      Width           =   435
   End
   Begin VB.CheckBox chkType 
      Caption         =   "2"
      Height          =   255
      Index           =   2
      Left            =   2760
      TabIndex        =   3
      Top             =   2280
      Width           =   435
   End
   Begin VB.CheckBox chkType 
      Caption         =   "1"
      Height          =   255
      Index           =   1
      Left            =   2160
      TabIndex        =   2
      Top             =   2280
      Width           =   435
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "PurcPRp06a.frx":0026
      Style           =   1  'Graphical
      TabIndex        =   27
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CheckBox optItm 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2160
      TabIndex        =   13
      Top             =   3240
      Width           =   735
   End
   Begin VB.CheckBox optExt 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2160
      TabIndex        =   12
      Top             =   3000
      Width           =   735
   End
   Begin VB.CheckBox optDsc 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2160
      TabIndex        =   11
      Top             =   2760
      Width           =   735
   End
   Begin VB.CheckBox optVnd 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2160
      TabIndex        =   10
      Top             =   2520
      Width           =   735
   End
   Begin VB.ComboBox txtDte 
      Height          =   315
      Left            =   2160
      TabIndex        =   1
      Tag             =   "4"
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6120
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6120
      TabIndex        =   15
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "PurcPRp06a.frx":07D4
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "PurcPRp06a.frx":0952
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.ComboBox cmbVnd 
      Height          =   315
      Left            =   2160
      TabIndex        =   0
      Tag             =   "3"
      Top             =   960
      Width           =   1555
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6120
      Top             =   1260
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4020
      FormDesignWidth =   7200
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Group Report By"
      Height          =   285
      Index           =   10
      Left            =   240
      TabIndex        =   31
      Top             =   3480
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Types"
      Height          =   285
      Index           =   9
      Left            =   240
      TabIndex        =   29
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label lblVEName 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   2160
      TabIndex        =   28
      Top             =   1320
      Width           =   3132
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Item Comments"
      Height          =   285
      Index           =   8
      Left            =   240
      TabIndex        =   26
      Top             =   3240
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Extended Descriptions"
      Height          =   285
      Index           =   7
      Left            =   240
      TabIndex        =   25
      Top             =   3000
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Descriptions"
      Height          =   285
      Index           =   6
      Left            =   240
      TabIndex        =   24
      Top             =   2760
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Vendor Information"
      Height          =   285
      Index           =   5
      Left            =   240
      TabIndex        =   23
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   288
      Index           =   4
      Left            =   240
      TabIndex        =   22
      Top             =   2040
      Width           =   1308
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   288
      Index           =   3
      Left            =   4200
      TabIndex        =   21
      Top             =   960
      Width           =   1308
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(On Or Before)"
      Height          =   288
      Index           =   2
      Left            =   4200
      TabIndex        =   20
      Top             =   1680
      Width           =   2028
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Items Due"
      Height          =   288
      Index           =   1
      Left            =   240
      TabIndex        =   19
      Top             =   1680
      Width           =   1308
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Vendor"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   14
      Top             =   960
      Width           =   1305
   End
End
Attribute VB_Name = "PurcPRp06a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
Option Explicit
Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd


Private Sub cmbVnd_Click()
   GetThisVendor
End Sub

Private Sub cmbVnd_LostFocus()
   cmbVnd = CheckLen(cmbVnd, 10)
   If Len(cmbVnd) = 0 Then cmbVnd = "ALL"
   GetThisVendor
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
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      AddComboStr cmbVnd.hwnd, "ALL"
      FillVendors
      cmbVnd = "ALL"
      GetThisVendor
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   
   GetOptions
   bOnLoad = 1
   txtDte = Format(ES_SYSDATE, "mm/dd/yyyy")
   
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveOptions
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set PurcPRp06a = Nothing
   
End Sub

Private Sub PrintReport()
    Dim sDate As String
    Dim sVendor As String
    Dim cCRViewer As EsCrystalRptViewer
    Dim sCustomReport As String
    Dim aRptPara As New Collection
    Dim aRptParaType As New Collection
    Dim aFormulaValue As New Collection
    Dim aFormulaName As New Collection
   
   sDate = Format(txtDte, "yyyy,mm,dd")
   sVendor = Compress(cmbVnd)
   
   On Error GoTo DiaErr1
   
 
    'get custom report name if one has been defined
    sCustomReport = GetCustomReport("prdpr06.rpt")

    Set cCRViewer = New EsCrystalRptViewer
    cCRViewer.Init
    cCRViewer.SetReportFileName sCustomReport, sReportPath

    cCRViewer.SetReportTitle = sCustomReport
    cCRViewer.ShowGroupTree False


   sSql = ""
   sSql = cCRViewer.GetReportSelectionFormula

   If (sSql <> "") Then sSql = sSql & " AND "
   sSql = sSql & "{PoitTable.PITYPE} = 14 AND " & _
            "{PoitTable.PIPDATE} <= " & CrystalDate(sDate)
   
'   sSql = "{PoitTable.PITYPE} = 14 AND {PoitTable.PIODDELIVERED} < 1 AND " & _
'            "{PoitTable.PIPDATE} <= " & CrystalDate(txtDte)
   If sVendor <> "ALL" Then
      sSql = sSql & " AND {VndrTable.VEREF}='" & sVendor & "'"
   End If
   
   'select part types
   Dim types As String
   Dim includes As String
   Dim I As Integer
   For I = 1 To 8
      If Me.chkType(I).Value = vbChecked Then
         If types = "" Then
            types = " AND ("
         Else
            types = types & " OR "
         End If
         types = types & "{PartTable.PALEVEL} = " & I
         includes = includes & " " & I
      End If
   Next
   If types = "" Then
      MsgBox "No part types selected"
      Exit Sub
   
   Else
      sSql = sSql & types & ")"
   End If
    
    ' set the report Selection
    cCRViewer.SetReportSelectionFormula (sSql)
   
    aFormulaName.Add "CompanyName"
    aFormulaName.Add "Includes"
    aFormulaName.Add "RequestBy"
    aFormulaName.Add "ShowVendor"
    aFormulaName.Add "ShowDescription"
    aFormulaName.Add "ShowExtendedDescription"
    aFormulaName.Add "ShowItem"
    aFormulaName.Add "GroupBy"
    

    aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
    aFormulaValue.Add CStr("'" & CStr(txtDte & " Types" & includes) & "'")
    aFormulaValue.Add CStr("'" & "Requested By:" & CStr(sInitials) & "'")
    aFormulaValue.Add optVnd.Value
    aFormulaValue.Add optDsc.Value
    aFormulaValue.Add optExt.Value
    aFormulaValue.Add optItm.Value
    aFormulaValue.Add CStr("'" & Me.cmbGroupBy(1).Text & "'")
    
    
    
    cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   
    cCRViewer.CRViewerSize Me
    ' Set report parameter
    cCRViewer.SetDbTableConnection
    ' print the copies
    cCRViewer.OpenCrystalReportObject Me, aFormulaName

    cCRViewer.ClearFieldCollection aRptPara
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

'BBS Removed this after searching the code. I can't find any reference to this routine anywhere.
'I didn't initially realize that and spent about 20 minutes researching the GROUPHDR section format syntax
'listed below only to find that this entire routine isn't even used. So, I figured I would remark this out
'to save the next developer the time I wasted.

'Private Sub PrintReport1()
'   Dim sDate As String
'   Dim sVendor As String
'   sDate = Format(txtDte, "yyyy,mm,dd")
'   sVendor = Compress(cmbVnd)
'
'   On Error GoTo DiaErr1
'   SetMdiReportsize MDISect
'   MDISect.Crw.Formulas(0) = "CompanyName='" & sFacility & "'"
'   'MDISect.Crw.Formulas(1) = "Includes=' " & txtDte & "...'"
'   MDISect.Crw.Formulas(2) = "RequestBy='Requested By: " & sInitials & "'"
'   sCustomReport = GetCustomReport("prdpr06")
'   MDISect.Crw.ReportFileName = sReportPath & sCustomReport
'   'sSql = "{PoitTable.PIPDATE}<=Date(" & sDate & ") "
'   sSql = "{PoitTable.PIPDATE} <= " & CrystalDate(txtDte) & vbCrLf
'   If sVendor <> "ALL" Then
'      sSql = sSql & "AND {VndrTable.VEREF}='" & sVendor & "'" & vbCrLf
'   End If
'
'   'select part types
'   Dim types As String
'   Dim includes As String
'   Dim i As Integer
'   For i = 1 To 8
'      If Me.chkType(i).Value = vbChecked Then
'         If types = "" Then
'            types = "AND ("
'         Else
'            types = types & "OR "
'         End If
'         types = types & "{PartTable.PALEVEL} = " & i & vbCrLf
'         includes = includes & " " & i
'      End If
'   Next
'   If types = "" Then
'      MsgBox "No part types selected"
'      Exit Sub
'
'   Else
'      sSql = sSql & types & ")"
'   End If
'
'   MDISect.Crw.Formulas(1) = "Includes='" & txtDte & " Types" & includes & "'"
'
'   If optVnd.Value = vbUnchecked Then
'      MDISect.Crw.SectionFormat(0) = "GROUPHDR.1.0;F;;;"
'      MDISect.Crw.SectionFormat(1) = "GROUPHDR.1.1;F;;;"
'      MDISect.Crw.SectionFormat(2) = "GROUPHDR.1.2;F;;;"
'   Else
'      MDISect.Crw.SectionFormat(0) = "GROUPHDR.1.0;T;;;"
'      MDISect.Crw.SectionFormat(1) = "GROUPHDR.1.1;T;;;"
'      MDISect.Crw.SectionFormat(2) = "GROUPHDR.1.2;T;;;"
'   End If
'   If optDsc.Value = vbUnchecked Then
'      MDISect.Crw.SectionFormat(3) = "GROUPFTR.4.0;F;;;"
'   Else
'      MDISect.Crw.SectionFormat(3) = "GROUPFTR.4.0;T;;;"
'   End If
'   If optExt.Value = vbUnchecked Then
'      MDISect.Crw.SectionFormat(4) = "GROUPFTR.3.0;F;;;"
'   Else
'      MDISect.Crw.SectionFormat(4) = "GROUPFTR.3.0;T;;;"
'   End If
'   If optItm.Value = vbUnchecked Then
'      MDISect.Crw.SectionFormat(5) = "GROUPFTR.2.0;F;;;"
'   Else
'      MDISect.Crw.SectionFormat(5) = "GROUPFTR.2.0;T;;;"
'   End If
'   MDISect.Crw.SelectionFormula = sSql
'   MouseCursor 13
'   SetCrystalAction Me
'   MouseCursor 0
'   Exit Sub
'
'DiaErr1:
'   sProcName = "printreport"
'   CurrError.Number = Err.Number
'   CurrError.Description = Err.Description
'   DoModuleErrors Me
'
'End Sub


Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub SaveOptions()
   Dim sOptions As String
   sOptions = optVnd & optDsc & optExt & optItm
   
   Dim I As Integer
   For I = 1 To 8
      sOptions = sOptions & chkType(I)
   Next
   sOptions = sOptions & LTrim(str(cmbGroupBy(1).ListIndex))   'BBS Added on 03/31/2010 for Ticket #27860
   
   
   SaveSetting "Esi2000", "EsiProd", "pr06", Trim(sOptions)
   
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   sOptions = GetSetting("Esi2000", "EsiProd", "pr06", sOptions) & "0000000000000"
   optVnd.Value = Left(sOptions, 1)
   optDsc.Value = Mid(sOptions, 2, 1)
   optExt.Value = Mid(sOptions, 3, 1)
   optItm.Value = Mid(sOptions, 4, 1)
   
   Dim I As Integer
   For I = 1 To 8
      chkType(I) = Mid(sOptions, I + 4, 1)
   Next
   If IsNumeric(Mid(sOptions, 13, 1)) Then cmbGroupBy(1).ListIndex = Val(Mid(sOptions, 13, 1)) Else cmbGroupBy(1).ListIndex = 0
   
End Sub

Private Sub optDis_Click()
   PrintReport
   
End Sub


Private Sub optDsc_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
End Sub


Private Sub optExt_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
End Sub


Private Sub optItm_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optPrn_Click()
   PrintReport
   
End Sub


Private Sub optVnd_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub txtDte_DropDown()
   ShowCalendarEx Me
   
End Sub


Private Sub txtDte_LostFocus()
   txtDte = CheckDateEx(txtDte)
   
End Sub
