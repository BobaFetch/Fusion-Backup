VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form Intavl03 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Part Availability For A Customer"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6300
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   6300
   Begin VB.CommandButton ShowPrinters 
      Height          =   250
      Left            =   0
      Picture         =   "Intavl03.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   25
      TabStop         =   0   'False
      ToolTipText     =   "Show Printers"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CheckBox optVew 
      Caption         =   "View"
      Height          =   255
      Left            =   3480
      TabIndex        =   24
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdFnd 
      DownPicture     =   "Intavl03.frx":018A
      Height          =   315
      Left            =   3840
      Picture         =   "Intavl03.frx":04CC
      Style           =   1  'Graphical
      TabIndex        =   23
      TabStop         =   0   'False
      ToolTipText     =   "Find A Part Number"
      Top             =   600
      Width           =   350
   End
   Begin VB.ComboBox cmbCst 
      Height          =   315
      Left            =   960
      TabIndex        =   2
      Tag             =   "3"
      Top             =   1320
      Width           =   1455
   End
   Begin VB.TextBox txtItm 
      Height          =   285
      Left            =   5160
      TabIndex        =   7
      Tag             =   "1"
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox cmbPrt 
      Height          =   285
      Left            =   960
      TabIndex        =   0
      Tag             =   "3"
      Top             =   585
      Width           =   2775
   End
   Begin VB.CheckBox optEDs 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2160
      TabIndex        =   4
      Top             =   2505
      Width           =   735
   End
   Begin VB.CheckBox optDsc 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2160
      TabIndex        =   3
      Top             =   2265
      Width           =   735
   End
   Begin VB.CheckBox optSls 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2160
      TabIndex        =   6
      Top             =   3000
      Width           =   735
   End
   Begin VB.CheckBox optLst 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2160
      TabIndex        =   5
      Top             =   2760
      Width           =   735
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   5040
      TabIndex        =   10
      Top             =   480
      Width           =   1335
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   720
         Picture         =   "Intavl03.frx":080E
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Print The Report"
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   120
         Picture         =   "Intavl03.frx":0998
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Display The Report"
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Close"
      Height          =   375
      Left            =   5160
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   0
      Width           =   1095
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5160
      Top             =   1680
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3570
      FormDesignWidth =   6300
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   960
      TabIndex        =   21
      Top             =   960
      Width           =   2775
   End
   Begin VB.Label lblNme 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   960
      TabIndex        =   20
      Top             =   1680
      Width           =   2775
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   255
      Index           =   7
      Left            =   2760
      TabIndex        =   19
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   18
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Number Of Items"
      Height          =   255
      Index           =   3
      Left            =   3840
      TabIndex        =   17
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales History"
      Height          =   255
      Index           =   6
      Left            =   360
      TabIndex        =   16
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Last Purchase"
      Height          =   255
      Index           =   5
      Left            =   360
      TabIndex        =   15
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Descriptions"
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   14
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ext. Descriptions"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   13
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   12
      Top             =   600
      Width           =   735
   End
   Begin VB.Label lblPrinter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Default Printer"
      ForeColor       =   &H00800000&
      Height          =   252
      Left            =   360
      TabIndex        =   11
      Top             =   0
      Width           =   2760
   End
End
Attribute VB_Name = "Intavl03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of                     ***
'*** ESI Software Engineering Inc, Seattle, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***

'See the UpdateTables prodecure for database revisions
' Intavl03 - Part Availiblity For A Customer (INTCOA Custom)
' Created: 06/19/03 (jcw)
' Revisions:
'8/19/03 (nth) Added Crystal Report
'6/30/05 fixed Part Lookup and Tab Order (per INTCOA)
'7/8/05 Request - return focus to cmbPrt after report
'5/9/08 Custom Report intavl03 removed from system (modified by INTCOA)
Option Explicit

Dim bOnLoad As Byte
Dim bCancel As Byte
Dim sMsg As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbCst_Click()
   FindCustomer Me, cmbCst
   
End Sub

Private Sub cmbCst_LostFocus()
   cmbCst = CheckLen(cmbCst, 10)
   If Len(cmbCst) Then
      FindCustomer Me, cmbCst
   Else
      cmbCst = "ALL"
   End If
   If cmbCst = "ALL" Then
      lblNme = "All Customers Selected."
   End If
   cmbPrt.TabIndex = 0
   
End Sub

Private Sub cmbPrt_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF4 Then
      ViewParts.lblControl = "CMBPRT"
      ViewParts.txtPrt = cmbPrt
      optVew.Value = vbChecked
      ViewParts.Show
   End If
   
End Sub

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, _
                             Shift As Integer, X As Single, Y As Single)
   bCancel = True
End Sub

Private Sub cmdVew_Click(Index As Integer)
   ViewParts.lblControl = "CMBPRT"
   ViewParts.txtPrt = cmbPrt
   ViewParts.Show
   
End Sub

Private Sub cmdFnd_Click()
   ViewParts.lblControl = "CMBPRT"
   ViewParts.txtPrt = cmbPrt
   optVew.Value = vbChecked
   ViewParts.Show
   
End Sub

Private Sub cmdHlp_Click()
   
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillCustomers
      FindCustomer Me, cmbCst
      bOnLoad = 0
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   
   GetOptions
   If optSls = vbUnchecked Then
      txtItm.Enabled = False
   End If
   bOnLoad = 1
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Private Sub SaveOptions()
   Dim sOptions As String
   sOptions = RTrim(optDsc.Value) _
              & RTrim(optEDs.Value) _
              & RTrim(optLst.Value) _
              & RTrim(optSls.Value)
   
   SaveSetting "Esi2000", "EsiFina", Me.Name, Trim(sOptions)
   SaveSetting "Esi2000", "EsiFina", Me.Name & "_Printer", lblPrinter
   SaveSetting "Esi2000", "EsiFina", Me.Name & "_Items", Trim(txtItm)
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiFina", Me.Name, sOptions)
   
   If Len(Trim(sOptions)) > 0 Then
      optDsc.Value = Val(Mid(sOptions, 1, 1))
      optEDs.Value = Val(Mid(sOptions, 2, 1))
      optLst.Value = Val(Mid(sOptions, 3, 1))
      optSls.Value = Val(Mid(sOptions, 4, 1))
   Else
      optDsc.Value = vbUnchecked
      optEDs.Value = vbUnchecked
      optLst.Value = vbUnchecked
      optSls.Value = vbUnchecked
   End If
   
   lblPrinter = GetSetting("Esi2000", "EsiFina", Me.Name & "_Printer", lblPrinter)
   txtItm = GetSetting("Esi2000", "EsiFina", Me.Name & "_Items", txtItm)
   
   If lblPrinter = "" Then lblPrinter = "Default Printer"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveOptions
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set Intavl03 = Nothing
End Sub

Private Sub optDis_Click()
   cmbPrt.TabIndex = 2
   PrintReport
   
End Sub

Private Sub optPrn_Click()
   PrintReport
End Sub

Private Sub optSls_Click()
   If optSls = vbChecked Then
      txtItm.Enabled = True
   Else
      txtItm.Enabled = False
   End If
End Sub

Private Sub cmbPrt_LostFocus()
   cmbPrt = CheckLen(cmbPrt, 30)
   If cmbPrt = "" Then cmbPrt = "ALL"
   
End Sub

Private Sub PrintReport()
   On Error GoTo DiaErr1
   
   Dim cCRViewer As EsCrystalRptViewer
    Dim sCustomReport As String
    Dim aRptPara As New Collection
    Dim aRptParaType As New Collection
    Dim aFormulaValue As New Collection
    Dim aFormulaName As New Collection
 
   
   If Trim(cmbPrt) = "" Then
      sMsg = "No Part Number Selected."
      MsgBox sMsg, vbExclamation, Caption
      cmbPrt.SetFocus
      Exit Sub
   End If
   
   MouseCursor 13
   optPrn.Enabled = False
   optDis.Enabled = False
   
   sCustomReport = GetCustomReport("intavl03")
    
    Set cCRViewer = New EsCrystalRptViewer
    cCRViewer.Init
    cCRViewer.SetReportFileName sCustomReport, sReportPath

    cCRViewer.SetReportTitle = sCustomReport
    cCRViewer.ShowGroupTree False
    
    aFormulaName.Add "CompanyName"
    aFormulaName.Add "Title1"
    aFormulaName.Add "Title2"
    aFormulaName.Add "RequestBy"
    aFormulaName.Add "Desc"
    aFormulaName.Add "ExtDesc"
    aFormulaName.Add "LastPO"
    
    aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
    aFormulaValue.Add CStr("'Part Availability For A Customer'")
    aFormulaValue.Add CStr("'Part Number: " & CStr(cmbPrt) & "  Customer: " & _
                                CStr(cmbCst) & "'")
    aFormulaValue.Add CStr("' Requested By: " & CStr(sInitials) & "'")
    aFormulaValue.Add CStr("'" & CStr(optDsc) & "'")
    aFormulaValue.Add CStr("'" & CStr(optEDs) & "'")
    aFormulaValue.Add CStr(CStr(optLst))
   
   If Val(txtItm) > 0 Then
        
    aFormulaName.Add "Item"
    aFormulaValue.Add CStr("'" & CStr(Val(txtItm)) & "'")
   
   End If
   
    aFormulaName.Add "Customer"
    aFormulaValue.Add CStr("'" & CStr(Compress(cmbCst)) & "'")
   
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   
   sSql = "{PartTable.PARTREF} = '" & Compress(cmbPrt) & "'"
   cCRViewer.SetReportSelectionFormula sSql
   
    cCRViewer.CRViewerSize Me
    
    ' Set report parameter
    cCRViewer.SetDbTableConnection


    cCRViewer.OpenCrystalReportObject Me, aFormulaName

    cCRViewer.ClearFieldCollection aRptPara
    cCRViewer.ClearFieldCollection aRptParaType
    cCRViewer.ClearFieldCollection aFormulaName
    cCRViewer.ClearFieldCollection aFormulaValue

      
   optPrn.Enabled = True
   optDis.Enabled = True
   MouseCursor 0
   ' cmbPrt.SetFocus
   Exit Sub
   
DiaErr1:
   On Error Resume Next
   optPrn.Enabled = True
   optDis.Enabled = True
   cmbPrt.SetFocus
   sProcName = "printreport"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

