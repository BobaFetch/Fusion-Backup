VERSION 5.00
Begin VB.Form PackPSp16a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Print Pack Slip Label"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7950
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   7950
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbEndPsn 
      Height          =   315
      Left            =   4560
      TabIndex        =   1
      ToolTipText     =   "Select Packing Slip (limit to 5000 desc) or type in a Packing Slip to repopulate drop-down"
      Top             =   1080
      Width           =   1575
   End
   Begin VB.ComboBox cmbPSItems 
      Height          =   315
      ItemData        =   "PackPSp16a.frx":0000
      Left            =   2520
      List            =   "PackPSp16a.frx":0022
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Tag             =   "9"
      ToolTipText     =   "Select the Pack Slip Item to Print"
      Top             =   1560
      Width           =   1335
   End
   Begin VB.ComboBox cmbLabelType 
      Height          =   315
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Tag             =   "9"
      Top             =   600
      Width           =   2415
   End
   Begin VB.TextBox txtCopies 
      Height          =   285
      Left            =   1680
      TabIndex        =   3
      Text            =   "1"
      Top             =   2040
      Width           =   345
   End
   Begin VB.ComboBox cmbPsn 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1680
      TabIndex        =   0
      ToolTipText     =   "Select Packing Slip (limit to 5000 desc) or type in a Packing Slip to repopulate drop-down"
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Close"
      Height          =   360
      Left            =   6840
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   120
      Width           =   1065
   End
   Begin VB.CommandButton optPrn 
      Height          =   330
      Left            =   7440
      Picture         =   "PackPSp16a.frx":0046
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Print The Report"
      Top             =   600
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.CommandButton optDis 
      Height          =   330
      Left            =   6840
      Picture         =   "PackPSp16a.frx":01D0
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Display The Report"
      Top             =   600
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.ComboBox cmbPrinters 
      Height          =   315
      Left            =   1680
      TabIndex        =   6
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label lblSelect 
      BackStyle       =   0  'Transparent
      Caption         =   "Packing Slip"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   15
      Top             =   1560
      Width           =   1065
   End
   Begin VB.Label lblSelect 
      BackStyle       =   0  'Transparent
      Caption         =   "Through"
      Height          =   285
      Index           =   3
      Left            =   3600
      TabIndex        =   14
      Top             =   1080
      Width           =   705
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label to Print"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   13
      Top             =   600
      Width           =   1425
   End
   Begin VB.Label lblSelect 
      BackStyle       =   0  'Transparent
      Caption         =   "Copies"
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   11
      Top             =   2040
      Width           =   1065
   End
   Begin VB.Label lblPrinter 
      Caption         =   "lblPrinter reqd by SetCrystalAction"
      Height          =   255
      Left            =   3720
      TabIndex        =   10
      Top             =   2280
      Visible         =   0   'False
      Width           =   3435
   End
   Begin VB.Label lblSelect 
      BackStyle       =   0  'Transparent
      Caption         =   "Packing Slip"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   9
      Top             =   1080
      Width           =   1065
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Available Printers"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   8
      Top             =   120
      Width           =   1785
   End
End
Attribute VB_Name = "PackPSp16a"
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
Dim sLabelPrinter As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown()  As New EsiKeyBd




Private Sub cmbLabelType_Click()
    If cmbLabelType.ListIndex <> 2 Then
        lblSelect(1).Caption = ""
        lblSelect(1).Visible = False
        'txtBoxes.Visible = True
        cmbPSItems.Visible = False
    Else
        cmbPSItems.Left = 1680
        lblSelect(1).Caption = "Pack Slip Item"
        lblSelect(1).Visible = True
        'txtBoxes.Visible = False
        cmbPSItems.Visible = True
        cmbPSItems.ListIndex = 0
        
    End If
End Sub

Private Sub cmbPrinters_Click()
    If sLabelPrinter <> cmbPrinters Then
        If MsgBox("Do you want to set this as your current shipping label printer?", vbYesNoCancel, "Shipping Label Printer") = vbYes Then sLabelPrinter = cmbPrinters
    End If
    
End Sub


Private Sub cmbPsn_Click()
    GetPackSlipInfo
    cmbEndPsn = cmbPsn
End Sub

Private Sub cmbPsn_LostFocus()
    GetPackSlipInfo
    'cmbEndPsn = cmbPsn
End Sub




Private Sub Form_Activate()
    Dim X As Printer
    On Error Resume Next
    MdiSect.BotPanel = Caption
    If bOnLoad Then
        For Each X In Printers
            If Left(X.DeviceName, 9) <> "Rendering" Then _
                cmbPrinters.AddItem X.DeviceName
        Next
        bOnLoad = 0
        FillPackingSlipCombo
        FillLabelCombo
        GetPackSlipInfo
        txtCopies.Text = "1"
        cmbLabelType.ListIndex = 0
        cmbEndPsn = cmbPsn
    End If
    MouseCursor 0
    
End Sub

Private Sub Form_Load()
    FormLoad Me
    FormatControls
    GetOptions
    bOnLoad = 1
End Sub

Private Sub cmdCan_Click()
    Unload Me

End Sub

Private Sub FormatControls()
    Dim b As Byte
    b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   ' txtBoxes = 1
    
End Sub


Private Sub GetOptions()
    Dim sOptions As String
    On Error Resume Next
    sLabelPrinter = GetSetting("Esi2000", "System", "PS Label Printer", sLabelPrinter)
    cmbPrinters = sLabelPrinter
    
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveSetting "Esi2000", "System", "PS Label Printer", sLabelPrinter
End Sub

Private Sub Form_Resize()
    Refresh
End Sub



Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set PackPSp16a = Nothing
End Sub



Private Sub optDis_Click()
   lblPrinter = cmbPrinters
   
   PrintLabels
End Sub

Private Sub optPrn_Click()
    lblPrinter = cmbPrinters
    
    PrintLabels

End Sub

Private Sub PrintLabels()
   Dim b           As Byte
   Dim FormDriver  As String
   Dim FormPort    As String
   Dim FormPrinter As String
   Dim iCopies As Integer
   Dim sRptFile As String
   
   If Len(cmbEndPsn) = 0 Then cmbEndPsn = cmbPsn
   If cmbPsn > cmbEndPsn Then
   MsgBox "Beginning Pack Slip Number cannot be Greater than Ending Pack Slip Number", vbOKOnly, Caption
    Exit Sub
   End If
   
   MouseCursor 13
   FormPrinter = Trim(sLabelPrinter)
   If Len(Trim(FormPrinter)) > 0 Then
      b = GetPrinterPort(FormPrinter, FormDriver, FormPort)
   Else
      FormPrinter = ""
      FormDriver = ""
      FormPort = ""
   End If
   iCopies = Val(txtCopies)
  
    MakeSureBoxRecordsExist cmbPsn, cmbEndPsn
   
    Dim cCRViewer As EsCrystalRptViewer
    Dim sCustomReport As String
    Dim aRptPara As New Collection
    Dim aRptParaType As New Collection
    Dim aFormulaValue As New Collection
    Dim aFormulaName As New Collection
    
    Select Case cmbLabelType.ListIndex
    Case 0: sRptFile = "sleps21.rpt"
    Case 1: sRptFile = "sleps23.rpt"
    Case Else
        sRptFile = "sleps22.rpt"
    End Select
    sCustomReport = GetCustomReport(sRptFile)

    Set cCRViewer = New EsCrystalRptViewer
    cCRViewer.Init
    cCRViewer.SetReportFileName sCustomReport, sReportPath
    cCRViewer.SetReportTitle = sRptFile
    cCRViewer.ShowGroupTree False

    
'    aFormulaName.Add "FromAdd1"
'    aFormulaValue.Add CStr("'" & CStr("U.S. CASTINGS LLC.") & "'")
    
    
    cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
    
    If cmbLabelType.ListIndex = 0 Or cmbLabelType.ListIndex = 1 Then
    
        sSql = "{PshdTable.PSNUMBER} IN '" & cmbPsn & "' TO '" & cmbEndPsn & "' "
            '"AND ( {PsibTable.PIBBOXNO} <= " & txtBoxes & " AND NOT(IsNull({PsibTable.PIBBOXNO})) )"
    Else
        sSql = "{PsitTable.PIPACKSLIP} IN '" & cmbPsn & "' TO '" & cmbEndPsn & "' " & _
            "AND {PsitTable.PIITNO}= " & cmbPSItems
    End If
    cCRViewer.SetReportSelectionFormula sSql
    cCRViewer.CRViewerSize Me
    cCRViewer.SetDbTableConnection
    ' print the copies
    cCRViewer.OpenCrystalReportObject Me, aFormulaName, iCopies

    cCRViewer.ClearFieldCollection aRptPara
    cCRViewer.ClearFieldCollection aFormulaName
    cCRViewer.ClearFieldCollection aFormulaValue
   
   MouseCursor 0
   Exit Sub

DiaErr1:
   sProcName = "PrintLabels"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me

End Sub



Private Sub FillPackingSlipCombo()
    cmbPsn.Clear

    On Error GoTo DiaErr1
    
    Dim RdoCmb As ADODB.Recordset

    sSql = "SELECT DISTINCT TOP(5000) PSNUMBER,PSTYPE,PIPACKSLIP,PSDATE" & vbCrLf _
      & "FROM PshdTable" & vbCrLf _
      & "JOIN PsitTable ON PSNUMBER=PIPACKSLIP" & vbCrLf _
      & "WHERE PSTYPE=1 AND PSSHIPPRINT=1 ORDER BY PSDATE DESC"
    bSqlRows = clsADOCon.GetDataSet(sSql, RdoCmb)
        If bSqlRows Then
            With RdoCmb
                cmbPsn = "" & Trim(!PsNumber)
                Do Until .EOF
                    AddComboStr cmbPsn.hwnd, "" & Trim(!PsNumber)
                    AddComboStr cmbEndPsn.hwnd, "" & Trim(!PsNumber)
                    .MoveNext
                Loop
                .Cancel
            End With
       End If
    Set RdoCmb = Nothing
    Exit Sub

DiaErr1:
    sProcName = "FillPackingSlipCombo"
    CurrError.Number = Err
    CurrError.Description = Err.Description
    DoModuleErrors Me
End Sub

Private Sub GetPackSlipInfo()
'    Dim rdoPS As ADODB.Recordset
'    Dim iNumBoxes As Integer
    
'    sSql = "SELECT PSBOXES FROM PshdTable WHERE PSNUMBER='" & cmbPsn & "' "
'
'    bSqlRows = clsADOCon.GetDataSet(sSql,rdoPS, ES_FORWARD)
'    If bSqlRows Then
'        iNumBoxes = Val(rdoPS!PSBOXES)
'        If iNumBoxes = 0 Then iNumBoxes = 1
'    Else
'        iNumBoxes = 1
'    End If
'    'UpDown1.Max = iNumBoxes
'    'txtBoxes.Text = LTrim(str(iNumBoxes))
'    Set rdoPS = Nothing
'
'    cmbPSItems.Clear
'    sSql = "SELECT DISTINCT PIITNO FROM PsitTable WHERE PIPACKSLIP='" & cmbPsn & "' ORDER BY PIITNO"
'    bSqlRows = clsADOCon.GetDataSet(sSql,rdoPS, ES_FORWARD)
'    If bSqlRows Then
'        While Not rdoPS.EOF
'            cmbPSItems.AddItem LTrim(str(rdoPS!PIITNO))
'            rdoPS.MoveNext
'        Wend
'
'    End If
'
'    Set rdoPS = Nothing
'    If cmbPSItems.ListCount > 0 Then cmbPSItems.ListIndex = 0
End Sub

'Private Sub txtBoxes_LostFocus()
'   UpdatePackSlipBoxes
'
'End Sub
'

'Private Sub UpDown1_Change()
'    UpdatePackSlipBoxes
'End Sub

'Private Sub UpdatePackSlipBoxes()
'    sSql = "UPDATE PshdTable SET PSBOXES=" & txtBoxes.Text & " WHERE PSNUMBER='" & cmbPsn & "' "
'    RdoCon.Execute sSql, rdExecDirect
'End Sub



Private Sub FillLabelCombo()
    Me.cmbLabelType.Clear
    cmbLabelType.AddItem "Pack Slip Label"
    cmbLabelType.AddItem "PACCAR Parts Label"
    If PrintingKanBanLabels Then cmbLabelType.AddItem "KANBAN Label"
    
End Sub
