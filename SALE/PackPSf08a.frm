VERSION 5.00
Begin VB.Form PackPSf08a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Print Pack Slip Label"
   ClientHeight    =   1710
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7050
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1710
   ScaleWidth      =   7050
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbASN 
      Height          =   315
      Left            =   1680
      TabIndex        =   7
      Top             =   960
      Width           =   2055
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Close"
      Height          =   360
      Left            =   5880
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   120
      Width           =   1065
   End
   Begin VB.CommandButton optPrn 
      Height          =   330
      Left            =   6480
      Picture         =   "PackPSf08a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Print The Report"
      Top             =   600
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.CommandButton optDis 
      Height          =   330
      Left            =   5880
      Picture         =   "PackPSf08a.frx":018A
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Display The Report"
      Top             =   600
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.ComboBox cmbPrinters 
      Height          =   315
      Left            =   1680
      TabIndex        =   2
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label lblPrinter 
      Caption         =   "lblPrinter reqd by SetCrystalAction"
      Height          =   255
      Left            =   2040
      TabIndex        =   6
      Top             =   1440
      Visible         =   0   'False
      Width           =   3435
   End
   Begin VB.Label lblSelect 
      BackStyle       =   0  'Transparent
      Caption         =   "ASN Number"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   960
      Width           =   1065
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Available Printers"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   1785
   End
End
Attribute VB_Name = "PackPSf08a"
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


Private Sub cmbPrinters_Click()
    If sLabelPrinter <> cmbPrinters Then
        If MsgBox("Do you want to set this as your current shipping label printer?", vbYesNoCancel, "Shipping Label Printer") = vbYes Then sLabelPrinter = cmbPrinters
    End If
    
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
        FillASNCombo
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
   Set PackPSf08a = Nothing
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
   
   MouseCursor 13
   FormPrinter = Trim(sLabelPrinter)
   If Len(Trim(FormPrinter)) > 0 Then
      b = GetPrinterPort(FormPrinter, FormDriver, FormPort)
   Else
      FormPrinter = ""
      FormDriver = ""
      FormPort = ""
   End If
   
    Dim cCRViewer As EsCrystalRptViewer
    Dim sCustomReport As String
    Dim aRptPara As New Collection
    Dim aRptParaType As New Collection
    Dim aFormulaValue As New Collection
    Dim aFormulaName As New Collection

 
    'get custom report name if one has been defined
    sCustomReport = GetCustomReport("sleps24.rpt")

 
    Set cCRViewer = New EsCrystalRptViewer
    cCRViewer.Init
    cCRViewer.SetReportFileName sCustomReport, sReportPath

    cCRViewer.SetReportTitle = "sleps24.rpt"
    cCRViewer.ShowGroupTree False
    
   aFormulaName.Add "CompanyName"
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
    
    cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
    
    sSql = "{PshdTable.PSCONTAINER}='" & cmbASN & "' "
    cCRViewer.SetReportSelectionFormula sSql
    cCRViewer.CRViewerSize Me
    cCRViewer.SetDbTableConnection
    ' print the copies
    cCRViewer.OpenCrystalReportObject Me, aFormulaName

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


Private Sub FillASNCombo()
    cmbASN.Clear

    On Error GoTo DiaErr1
    
    Dim RdoCmb As ADODB.Recordset

    sSql = "SELECT DISTINCT TOP(5000) PSCONTAINER " & vbCrLf _
      & "FROM PsitTable " & vbCrLf _
      & "INNER JOIN PshdTable ON PSNUMBER=PIPACKSLIP " & vbCrLf _
      & "WHERE PSTYPE=1 AND PSSHIPPRINT=1 AND LEN(PSCONTAINER)>0 ORDER BY PSCONTAINER DESC"
    Debug.Print sSql
    
    bSqlRows = clsADOCon.GetDataSet(sSql, RdoCmb)
        If bSqlRows Then
            With RdoCmb
                cmbASN = "" & Trim(!PSCONTAINER)
                Do Until .EOF
                    AddComboStr cmbASN.hwnd, "" & Trim(!PSCONTAINER)
                    .MoveNext
                Loop
                .Cancel
            End With
       End If
    Set RdoCmb = Nothing
    Exit Sub

DiaErr1:
    sProcName = "FillASNCombo"
    CurrError.Number = Err
    CurrError.Description = Err.Description
    DoModuleErrors Me
End Sub



