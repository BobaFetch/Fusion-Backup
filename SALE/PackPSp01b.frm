VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form PackPSp01b 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Inventory Labels (Prepackaging)"
   ClientHeight    =   2220
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   8085
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2220
   ScaleWidth      =   8085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport CRWLabels 
      Left            =   480
      Top             =   2040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.TextBox txtQty 
      Height          =   288
      Left            =   7500
      TabIndex        =   2
      Tag             =   "2"
      Text            =   "1"
      ToolTipText     =   "Number Of Labels"
      Top             =   1680
      Width           =   372
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Height          =   315
      Left            =   5280
      TabIndex        =   1
      ToolTipText     =   "Save This Printer As The Label Printer"
      Top             =   120
      Width           =   915
   End
   Begin VB.ComboBox lblPrinter 
      Height          =   288
      Left            =   1920
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "PackPSp01b.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   250
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Close"
      Height          =   360
      Left            =   6840
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6840
      TabIndex        =   4
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "PackPSp01b.frx":07AE
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   600
         Picture         =   "PackPSp01b.frx":092C
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   120
      Top             =   3000
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2220
      FormDesignWidth =   8085
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Packing Slip"
      Height          =   285
      Index           =   6
      Left            =   300
      TabIndex        =   19
      Top             =   960
      Width           =   1785
   End
   Begin VB.Label lblPackingSlipNumber 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1860
      TabIndex        =   18
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Item"
      Height          =   285
      Index           =   5
      Left            =   4380
      TabIndex        =   17
      Top             =   960
      Width           =   645
   End
   Begin VB.Label lblPackingSlipItem 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   5220
      TabIndex        =   16
      Top             =   960
      Width           =   855
   End
   Begin VB.Label lblQuantity 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   5220
      TabIndex        =   15
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity"
      Height          =   285
      Index           =   4
      Left            =   4380
      TabIndex        =   14
      Top             =   1680
      Width           =   1785
   End
   Begin VB.Label lblPartNumber 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1860
      TabIndex        =   13
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      Height          =   285
      Index           =   3
      Left            =   300
      TabIndex        =   12
      Top             =   1680
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Order"
      Height          =   285
      Index           =   2
      Left            =   300
      TabIndex        =   11
      Top             =   1320
      Width           =   1785
   End
   Begin VB.Label lblSalesOrder 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1860
      TabIndex        =   10
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label Quantity"
      Height          =   285
      Index           =   1
      Left            =   6300
      TabIndex        =   9
      Top             =   1680
      Width           =   1185
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Available Printers"
      Height          =   288
      Index           =   0
      Left            =   360
      TabIndex        =   3
      Top             =   120
      Width           =   1788
   End
End
Attribute VB_Name = "PackPSp01b"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
'4/23/07 CJS New
Option Explicit
Dim bOnLoad As Byte

Dim sLabelPrinter As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd




Private Sub cmdApply_Click()
   sLabelPrinter = lblPrinter
   If Len(Trim(lblPrinter)) Then
      MsgBox lblPrinter & " Saved As Your Label Printer.", _
         vbInformation, Caption
   Else
      MsgBox "No Printer was Saved As Your Label Printer.", _
         vbInformation, Caption
   End If
   
End Sub

Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdHlp_Click()
   '    If cmdHlp Then
   '        MouseCursor 13
   '        OpenHelpContext 907
   '        MouseCursor 0
   '        cmdHlp = False
   '    End If
   
End Sub




Private Sub Form_Activate()
   Dim X As Printer
   On Error Resume Next
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      For Each X In Printers
         If Left(X.DeviceName, 9) <> "Rendering" Then _
                 lblPrinter.AddItem X.DeviceName
      Next
      bOnLoad = 0
      
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
   SaveSetting "Esi2000", "System", "Label Printer", sLabelPrinter
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set PackPSp01b = Nothing
   
End Sub

Private Sub PrintLabels()
    Dim b As Byte
    Dim FormDriver As String
    Dim FormPort As String
    Dim FormPrinter As String
   
   
    MouseCursor 13
    'SetMdiReportsize MdiSect
   
    Dim cCRViewer As EsCrystalRptViewer
    Dim sCustomReport As String
    Dim aRptPara As New Collection
    Dim aRptParaType As New Collection
    Dim aFormulaValue As New Collection
    Dim aFormulaName As New Collection
 
   
    'get custom report name if one has been defined
    sCustomReport = GetCustomReport("sleps20.rpt")
 
    Set cCRViewer = New EsCrystalRptViewer
    cCRViewer.Init
    cCRViewer.SetReportFileName sCustomReport, sReportPath

    cCRViewer.SetReportTitle = "sleps20.rpt"
    cCRViewer.ShowGroupTree False

    aFormulaName.Add "PartNumber"
    aFormulaName.Add "PartRef"
    aFormulaName.Add "Quantity"
    aFormulaName.Add "PackSlipNumber"
    aFormulaName.Add "PackSlipItem"

    aFormulaValue.Add CStr("'" & CStr(lblPartNumber) & "'")
    aFormulaValue.Add CStr("'" & CStr(Compress(lblPartNumber)) & "'")
    aFormulaValue.Add CStr("'" & CStr(lblQuantity) & "'")
    aFormulaValue.Add CStr("'" & CStr(lblPackingSlipNumber) & "'")
    aFormulaValue.Add CStr("'" & CStr(lblPackingSlipItem) & "'")

    cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
    
    cCRViewer.CRViewerSize Me
    ' Set report parameter
    cCRViewer.SetDbTableConnection
    ' print the copies
    cCRViewer.OpenCrystalReportObject Me, aFormulaName, Val(txtQty)

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
    MouseCursor ccArrow
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtQty = 1
   
End Sub


Private Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   sLabelPrinter = GetSetting("Esi2000", "System", "Label Printer", sLabelPrinter)
   lblPrinter = sLabelPrinter
   
End Sub

Private Sub optDis_Click()
   PrintLabels
End Sub

Private Sub optPrn_Click()
   PrintLabels
End Sub

Private Sub txtQty_LostFocus()
   txtQty = Abs(Val(txtQty))
   If Val(txtQty) < 1 Then txtQty = 1
   
End Sub
