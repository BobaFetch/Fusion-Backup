VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form ShopSHp14a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Manufacturing Orders Splits"
   ClientHeight    =   3060
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7260
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3060
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "ShopSHp14a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   17
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox txtEnd 
      Height          =   315
      Left            =   4200
      TabIndex        =   2
      Tag             =   "4"
      Top             =   1560
      Width           =   1250
   End
   Begin VB.ComboBox txtBeg 
      Height          =   315
      Left            =   2040
      TabIndex        =   1
      Tag             =   "4"
      Top             =   1560
      Width           =   1250
   End
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   2040
      TabIndex        =   0
      Top             =   1080
      Width           =   3255
   End
   Begin VB.CheckBox optExt 
      Caption         =   "____"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   2040
      TabIndex        =   4
      Top             =   2400
      Width           =   735
   End
   Begin VB.CheckBox optDsc 
      Caption         =   "____"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   2040
      TabIndex        =   3
      Top             =   2160
      Width           =   735
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6120
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6120
      TabIndex        =   6
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "ShopSHp14a.frx":07AE
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "ShopSHp14a.frx":092C
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6840
      Top             =   2640
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3060
      FormDesignWidth =   7260
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   6
      Left            =   5520
      TabIndex        =   16
      Top             =   1560
      Width           =   1425
   End
   Begin VB.Label P 
      BackStyle       =   0  'Transparent
      Caption         =   "Extended Descriptions"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   15
      Top             =   2400
      Width           =   1785
   End
   Begin VB.Label z1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Through"
      Height          =   285
      Index           =   4
      Left            =   3120
      TabIndex        =   14
      Top             =   1560
      Width           =   1305
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Split From"
      Height          =   285
      Index           =   3
      Left            =   240
      TabIndex        =   13
      Top             =   1560
      Width           =   2145
   End
   Begin VB.Label P 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Descriptions"
      Height          =   285
      Index           =   3
      Left            =   240
      TabIndex        =   12
      Top             =   2160
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   2
      Left            =   5520
      TabIndex        =   11
      Top             =   1080
      Width           =   1545
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   10
      Tag             =   " "
      Top             =   1920
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Manufacturing Order(s)"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   1080
      Width           =   2145
   End
End
Attribute VB_Name = "ShopSHp14a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
Option Explicit
Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub FillCombo()
   On Error GoTo DiaErr1
   sSql = "SELECT DISTINCT SPLIT_FROMRUNREF,PARTREF,PARTNUM FROM " _
          & "RnspTable,PartTable WHERE SPLIT_FROMRUNREF=PARTREF " _
          & "ORDER BY SPLIT_FROMRUNREF"
   LoadComboBox cmbPrt, 1
   If cmbPrt.ListCount > 0 Then cmbPrt = cmbPrt.List(0)
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub



Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      'SelectHelpTopic Me, "Reports"
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub



Private Sub Form_Activate()
   On Error Resume Next
   MDISect.lblBotPanel = Caption
   If bOnLoad Then FillCombo
   bOnLoad = 0
   MouseCursor 0
   
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
   Set ShopSHp14a = Nothing
   
End Sub
Private Sub PrintReport()

   Dim sBegDate As String
   Dim sEndDate As String
   Dim sMans As String
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection

   MouseCursor 13
   
   On Error GoTo DiaErr1
   
   If Len(Trim(txtBeg)) = 0 Then txtBeg = "ALL"
   If Len(Trim(txtEnd)) = 0 Then txtEnd = "ALL"
   If Not IsDate(txtBeg) Then
      sBegDate = "1995,01,01"
   Else
      sBegDate = Format(txtBeg, "yyyy,mm,dd")
   End If
   If Not IsDate(txtEnd) Then
      sEndDate = "2024,12,31"
   Else
      sEndDate = Format(txtEnd, "yyyy,mm,dd")
   End If
   
   If cmbPrt = "" Then cmbPrt = "ALL"
   If cmbPrt <> "ALL" Then sMans = Compress(cmbPrt)
    
    Set cCRViewer = New EsCrystalRptViewer
    cCRViewer.Init
    sCustomReport = GetCustomReport("prdsh20")
    cCRViewer.SetReportFileName sCustomReport, sReportPath
    cCRViewer.SetReportTitle = sCustomReport

    aFormulaName.Add "CompanyName"
    aFormulaName.Add "RequestBy"
    aFormulaName.Add "Includes"
    aFormulaName.Add "ShowDescription"
    aFormulaName.Add "ShowExtendedDescription"

    aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
    aFormulaValue.Add CStr("'Requested By:" & CStr(sInitials) & "'")
    aFormulaValue.Add CStr("'" & cmbPrt & " From " _
                        & txtBeg & " To " & txtEnd & "...'")
    aFormulaValue.Add optDsc.Value
    aFormulaValue.Add optExt.Value
    
    cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
    
 '  MDISect.Crw.ReportFileName = sReportPath & sCustomReport
   sSql = "{RnspTable.SPLIT_FROMRUNREF} LIKE '" & sMans & "*' AND " _
          & "{RnspTable.SPLIT_SPLDATE} in Date(" & Format(sBegDate, "yyyy,mm,dd") & ") " _
          & "to Date(" & Format(sEndDate, "yyyy,mm,dd") & ")"
   
'   If optDsc.value = vbUnchecked Then
'      MDISect.Crw.SectionFormat(0) = "GROUPHDR.0.1;F;;;"
'   Else
'      MDISect.Crw.SectionFormat(0) = "GROUPHDR.0.1;T;;;"
'   End If
'   If optExt.value = vbUnchecked Then
'      MDISect.Crw.SectionFormat(1) = "GROUPHDR.0.2;F;;;"
'   Else
'      MDISect.Crw.SectionFormat(1) = "GROUPHDR.0.2;T;;;"
'   End If
'   MDISect.Crw.SelectionFormula = sSql
'   SetCrystalAction Me
   
   cCRViewer.SetReportSelectionFormula sSql
   cCRViewer.SetDbTableConnection
   
   cCRViewer.ShowGroupTree False
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


Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
'   txtEnd = Format(Now, "mm/dd/yy")
'   txtBeg = Left(txtEnd, 3) & "01" & Right(txtEnd, 3)
   txtEnd = ""
   txtBeg = ""
   
End Sub

Private Sub SaveOptions()
   Dim sOptions As String
   sOptions = Trim$(str$(optDsc.Value))
   sOptions = sOptions & Trim$(str$(optExt.Value))
   SaveSetting "Esi2000", "EsiProd", "sh20", sOptions
   
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiProd", "sh20", sOptions)
   If Trim(sOptions) <> "" Then
      optDsc.Value = Val(Left(sOptions, 1))
      optExt.Value = Val(Right(sOptions, 1))
   Else
      optDsc.Value = vbChecked
      optExt.Value = vbChecked
   End If
   
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


Private Sub optPrn_Click()
   PrintReport
   
End Sub


Private Sub txtBeg_DropDown()
   ShowCalendarEx Me
   
End Sub


Private Sub txtBeg_LostFocus()
   If Trim(txtBeg) <> "" And Trim(txtBeg) <> "ALL" Then
      txtBeg = CheckDateEx(txtBeg)
   Else
      txtBeg = "ALL"
   End If
   
End Sub


Private Sub txtend_DropDown()
   ShowCalendarEx Me
   
End Sub


Private Sub txtEnd_LostFocus()
   If Trim(txtEnd) <> "" And Trim(txtEnd) <> "ALL" Then
      txtEnd = CheckDateEx(txtEnd)
   Else
      txtEnd = "ALL"
   End If
End Sub
