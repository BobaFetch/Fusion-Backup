VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form BompBMp05a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "List of Revised Parts Lists"
   ClientHeight    =   3360
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7230
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3360
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   2160
      TabIndex        =   0
      Top             =   1080
      Width           =   3075
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "BompBMp05a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   21
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CheckBox optExt 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   285
      Left            =   2160
      TabIndex        =   5
      Top             =   2520
      Value           =   1  'Checked
      Width           =   735
   End
   Begin VB.CheckBox optDsc 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2160
      TabIndex        =   4
      Top             =   2280
      Value           =   1  'Checked
      Width           =   735
   End
   Begin VB.ComboBox txtBeg 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   2160
      TabIndex        =   2
      Tag             =   "4"
      Top             =   1560
      Width           =   1250
   End
   Begin VB.ComboBox txtEnd 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   4200
      TabIndex        =   3
      Tag             =   "4"
      Top             =   1560
      Width           =   1250
   End
   Begin VB.TextBox txtPrt 
      Height          =   285
      Left            =   2160
      TabIndex        =   1
      Tag             =   "3"
      Text            =   "ALL"
      Top             =   1110
      Visible         =   0   'False
      Width           =   3075
   End
   Begin VB.CommandButton cmdFnd 
      DownPicture     =   "BompBMp05a.frx":07AE
      Height          =   315
      Left            =   5280
      Picture         =   "BompBMp05a.frx":0AF0
      Style           =   1  'Graphical
      TabIndex        =   14
      TabStop         =   0   'False
      ToolTipText     =   "Find A Part Number"
      Top             =   1110
      Visible         =   0   'False
      Width           =   350
   End
   Begin VB.CheckBox optDet 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2160
      TabIndex        =   6
      Top             =   2840
      Width           =   735
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6120
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6120
      TabIndex        =   7
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "BompBMp05a.frx":0E32
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "BompBMp05a.frx":0FB0
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   120
      Top             =   3600
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3360
      FormDesignWidth =   7230
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   288
      Index           =   9
      Left            =   5880
      TabIndex        =   20
      Top             =   1560
      Width           =   1404
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Descriptions"
      Height          =   285
      Index           =   7
      Left            =   240
      TabIndex        =   19
      Top             =   2280
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Extended Descriptions"
      Height          =   285
      Index           =   6
      Left            =   240
      TabIndex        =   18
      Top             =   2535
      Width           =   1815
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Revised Dates From"
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   17
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Through"
      Height          =   255
      Index           =   4
      Left            =   3480
      TabIndex        =   16
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number(s)"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   15
      Top             =   1110
      Width           =   1815
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Type 4 Lists"
      Height          =   285
      Index           =   3
      Left            =   240
      TabIndex        =   13
      Top             =   2840
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   2
      Left            =   5880
      TabIndex        =   12
      Top             =   1110
      Width           =   1400
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   11
      Tag             =   " "
      Top             =   2040
      Width           =   1425
   End
End
Attribute VB_Name = "BompBMp05a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
'5/7/05 New
Option Explicit
Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub FillCombo()
   '   Dim RdoCmb As rdoResultset
   '   On Error GoTo DiaErr1
   '    sSql = "SELECT DISTINCT CUREF,CUNICKNAME,SOCUST FROM " _
   '        & "CustTable,SohdTable WHERE CUREF=SOCUST"
   '    bSqlRows = GetDataSet(RdoCmb, ES_FORWARD)
   '        If bSqlRows Then
   '            With RdoCmb
   '                Do Until .EOF
   '                    AddComboStr cmbCst.hWnd, "" & Trim(!CUNICKNAME)
   '                    .MoveNext
   '                Loop
   '                .Cancel
   '            End With
   '        Else
   '            lblNme = "*** No Customers With SO's Found ***"
   '        End If
   '    Set RdoCmb = Nothing
   '    cmbCst = "ALL"
   '    GetCustomer
   '   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdFnd_Click()
   ViewParts.lblControl = "TXTPRT"
   ViewParts.txtPrt = txtPrt
   ViewParts.Show
   
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
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      bOnLoad = 0
      
      Dim bPartSearch As Boolean
      
      bPartSearch = GetPartSearchOption
      SetPartSearchOption (bPartSearch)
      
      If (Not bPartSearch) Then FillPartCombo cmbPrt
      
      cmbPrt = "ALL"
   End If
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
   Set BompBMp05a = Nothing
   
End Sub

Private Sub PrintReport()
   Dim sPart As String
   Dim sBegDate As String
   Dim sEndDate As String
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   
   MouseCursor 13
   
   On Error GoTo DiaErr1
   If cmbPrt = "" Then cmbPrt = "ALL"
   If cmbPrt <> "ALL" Then sPart = Compress(cmbPrt)
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
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "Includes"
   aFormulaName.Add "RequestBy"
   aFormulaName.Add "ShowDesc"
   aFormulaName.Add "ShowExDesc"
   
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'Part Number(s) " & CStr(cmbPrt & " From " _
                        & txtBeg & " Through " & txtEnd) & "'")
   aFormulaValue.Add CStr("'Requested By:" & CStr(sInitials) & "'")
   aFormulaValue.Add optDsc.Value
   aFormulaValue.Add OptExt.Value
   
   sCustomReport = GetCustomReport("engbm05")
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue

   sSql = "{BmhdTable.BMHREF} LIKE '" & sPart & "*' AND " _
          & "{BmhdTable.BMHREVDATE} in Date(" & sBegDate & ") " _
          & "to Date(" & sEndDate & ") "
   If optDet.Value = vbUnchecked Then
      sSql = sSql & "AND {PartTable.PALEVEL}<4"
   Else
      sSql = sSql & "AND {PartTable.PALEVEL}<5"
   End If
   cCRViewer.SetReportSelectionFormula sSql
   cCRViewer.CRViewerSize Me
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
   txtPrt = "ALL"
 '  txtEnd = Format(Now, "mm/dd/yy")
 '  txtBeg = Left(txtEnd, 3) & "01" & Right(txtEnd, 3)
   txtEnd = ""
   txtBeg = ""
End Sub

Private Sub SaveOptions()
   Dim sOptions As String
   On Error Resume Next
   sOptions = Trim(str(optDsc.Value)) _
              & Trim(str(OptExt.Value)) _
              & Trim(str(optDet.Value))
   SaveSetting "Esi2000", "EsiEngr", "pl05", sOptions
   
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiEngr", "pl05", sOptions)
   If Len(Trim(sOptions)) > 0 Then
      optDsc.Value = Val(Left(sOptions, 1))
      OptExt.Value = Val(Mid(sOptions, 2, 1))
      optDet.Value = Val(Right(sOptions, 1))
   End If
   
End Sub

Private Sub optDet_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
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
   If Len(Trim(txtBeg)) = 0 Then txtBeg = "ALL"
   If txtBeg <> "ALL" Then txtBeg = CheckDateEx(txtBeg)
   
End Sub

Private Sub txtEnd_DropDown()
   ShowCalendarEx Me
   
End Sub

Private Sub txtEnd_LostFocus()
   If Len(Trim(txtEnd)) = 0 Then txtEnd = "ALL"
   If txtEnd <> "ALL" Then txtEnd = CheckDateEx(txtEnd)
   
End Sub

Private Sub txtPrt_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF4 Then
      ViewParts.lblControl = "TXTPRT"
      ViewParts.txtPrt = txtPrt
      ViewParts.Show
   End If
   
End Sub

Private Sub cmbPrt_LostFocus()
   cmbPrt = CheckLen(cmbPrt, 30)
   If Len(cmbPrt) = 0 Then cmbPrt = "ALL"
   
End Sub

Function SetPartSearchOption(bPartSearch As Boolean)
   
   If (bPartSearch = True) Then
      cmbPrt.Visible = False
      txtPrt.Visible = True
      cmdFnd.Visible = True
   Else
      cmbPrt.Visible = True
      txtPrt.Visible = False
      cmdFnd.Visible = False
   End If
End Function

Private Sub txtPrt_LostFocus()
   txtPrt = CheckLen(txtPrt, 30)
   If Trim(txtPrt) = "" Or Trim(txtPrt) = "ALL" Then txtPrt = "ALL"
   cmbPrt = txtPrt
End Sub

