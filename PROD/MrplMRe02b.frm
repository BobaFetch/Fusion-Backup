VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Begin VB.Form MrplMRe02b 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PO Comments"
   ClientHeight    =   3045
   ClientLeft      =   2985
   ClientTop       =   2310
   ClientWidth     =   7980
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3045
   ScaleWidth      =   7980
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton optPrn 
      Height          =   330
      Left            =   6120
      Picture         =   "MrplMRe02b.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Print The Report"
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.TextBox txtTotItems 
      Enabled         =   0   'False
      Height          =   375
      Left            =   3960
      TabIndex        =   10
      Top             =   240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ComboBox cmbItem 
      Height          =   315
      ItemData        =   "MrplMRe02b.frx":018A
      Left            =   4680
      List            =   "MrplMRe02b.frx":018C
      Sorted          =   -1  'True
      TabIndex        =   8
      Tag             =   "2"
      ToolTipText     =   "Engineer"
      Top             =   840
      Width           =   1380
   End
   Begin VB.TextBox txtPONumber 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1560
      TabIndex        =   7
      Top             =   240
      Width           =   2055
   End
   Begin VB.ComboBox cmbComType 
      Height          =   315
      ItemData        =   "MrplMRe02b.frx":018E
      Left            =   1560
      List            =   "MrplMRe02b.frx":0190
      Sorted          =   -1  'True
      TabIndex        =   4
      Tag             =   "2"
      ToolTipText     =   "Engineer"
      Top             =   840
      Width           =   1740
   End
   Begin VB.TextBox txtCmt 
      Height          =   1185
      Left            =   1560
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Tag             =   "9"
      Text            =   "MrplMRe02b.frx":0192
      ToolTipText     =   "Comment (5120 Chars Max)"
      Top             =   1440
      Width           =   4575
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   285
      Left            =   6360
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Fill The Current Operation"
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "&Done"
      Height          =   315
      Left            =   6840
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   990
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6240
      Top             =   3240
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3045
      FormDesignWidth =   7980
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "PO Item"
      Height          =   285
      Index           =   1
      Left            =   3960
      TabIndex        =   9
      Top             =   840
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "PO Number"
      Height          =   285
      Index           =   0
      Left            =   360
      TabIndex        =   6
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "CommentsType"
      Height          =   285
      Index           =   2
      Left            =   360
      TabIndex        =   5
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Comments"
      Height          =   285
      Index           =   14
      Left            =   360
      TabIndex        =   3
      Top             =   1440
      Width           =   945
   End
End
Attribute VB_Name = "MrplMRe02b"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'11/2/04 Revised general structure and Fill button
'        Attempts to update Ops grid
'1/26/07 Undo
Option Explicit

Dim bOnLoad As Byte
Dim iUserLogo As Integer

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub



Private Sub cmbComType_Click()
   If (cmbComType = "PO Remarks") Then
      cmbItem = ""
      cmbItem.Enabled = False
   Else
      cmbItem.Enabled = True
   End If
End Sub


Private Sub cmdAdd_Click()
   
   Dim strPONum  As String
   Dim strPOItem As String
   
   On Error GoTo DiaErr1
   strPONum = txtPONumber
   If (strPONum <> "") Then
   
      If (cmbComType = "PO Remarks") Then
         txtCmt = CheckLen(txtCmt, 6000)
         sSql = "UPDATE PohdTable SET POREMARKS='" & txtCmt & "' WHERE PONUMBER = " & Val(strPONum)
         clsADOCon.ExecuteSql sSql
      Else
         strPOItem = cmbItem
         If (strPOItem <> "") Then
            txtCmt = CheckLen(txtCmt, 2048)
            txtCmt = StrCase(txtCmt, ES_FIRSTWORD)
            
            sSql = "UPDATE PoitTable SET PICOMT = '" & txtCmt & "' " _
               & " WHERE PINUMBER= " & Val(strPONum) & " AND PIITEM = " & Val(strPOItem)
      
            clsADOCon.ExecuteSql sSql
         Else
            MsgBox "Please select a PO Item.", _
                        vbExclamation, Caption

         End If
      End If
   End If
   
   Exit Sub
   
DiaErr1:
   sProcName = "Add Comment Failed"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me

End Sub

Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub FillCommentType()
   'Dim RdoCmb As ADODB.Recordset
   On Error GoTo DiaErr1
   cmbComType.AddItem ("PO Remarks"), 0
   cmbComType.AddItem ("PO Item Comments"), 1
   If cmbComType.ListCount > 0 Then cmbComType = cmbComType.List(0)
   Exit Sub
   
DiaErr1:
   sProcName = "FillCommentType"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub FillPoItems(iTotItem As Integer)
   'Dim RdoCmb As ADODB.Recordset
   On Error GoTo DiaErr1
   Dim I As Integer
   cmbItem.AddItem ("")
   For I = 1 To iTotItem
      cmbItem.AddItem (I)
   Next
   If cmbItem.ListCount > 0 Then cmbItem = cmbItem.List(0)
   Exit Sub
   
DiaErr1:
   sProcName = "FillPoItems"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub



Private Sub Form_Activate()
   MouseCursor 13
   If bOnLoad Then
      
      FillCommentType
      FillPoItems txtTotItems
      cmbItem.Enabled = False
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   SetFormSize Me
   Move 2000, 2000
   FormatControls
   bOnLoad = 1
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   MouseCursor 0
   Set MrplMRe02b = Nothing
   
End Sub

Private Sub optPrn_Click()
   PrintReport
End Sub

Private Sub PrintReport()
   Dim bForm As Byte
   clsADOCon.ExecuteSql "UPDATE ComnTable SET CURPONUMBER=" & txtPONumber & " "
   bForm = GetPrintedForm("purchase order")
   On Error GoTo Ppr01
   
   Dim cCRViewer As EsCrystalRptViewer
    Dim sCustomReport As String
    Dim aRptPara As New Collection
    Dim aRptParaType As New Collection
    Dim aFormulaValue As New Collection
    Dim aFormulaName As New Collection
   
    Set cCRViewer = New EsCrystalRptViewer
    cCRViewer.Init
    
    sCustomReport = GetCustomReport("prdpr01")
    cCRViewer.SetReportFileName sCustomReport, sReportPath

    cCRViewer.SetReportTitle = sCustomReport
    cCRViewer.ShowGroupTree False

   
    aFormulaName.Add "Company"
    aFormulaName.Add "Phone"
    aFormulaName.Add "Fax"
    aFormulaName.Add "ShowComt"
    aFormulaName.Add "ShowRem"
    aFormulaName.Add "CoAddress1"
    aFormulaName.Add "CoAddress2"
    aFormulaName.Add "CoAddress3"
    aFormulaName.Add "CoAddress4"
    aFormulaName.Add "PoNumber"
    aFormulaName.Add "ShowExtDesc"
    aFormulaName.Add "ShowCanceledItems"
    aFormulaName.Add "ShowMoAllocations"
    aFormulaName.Add "ShowOpComments"
    aFormulaName.Add "ShowRecInfo"
    aFormulaName.Add "ShowServPartDoc"
    aFormulaName.Add "ShowPartDoc"
    aFormulaName.Add "ShowOurAddress"
    aFormulaName.Add "ResaleNumber"
    
    
    aFormulaValue.Add CStr("'" & CStr(Co.Name) & "'")
    aFormulaValue.Add CStr("'" & CStr(Co.Phone) & "'")
    aFormulaValue.Add CStr("'" & CStr(Co.Fax) & "'")
    aFormulaValue.Add CStr("'1'")
    aFormulaValue.Add CStr("'1'")
    aFormulaValue.Add CStr("'" & CStr(Co.Addr(1)) & "'")
    aFormulaValue.Add CStr("'" & CStr(Co.Addr(2)) & "'")
    aFormulaValue.Add CStr("'" & CStr(Co.Addr(3)) & "'")
    aFormulaValue.Add CStr("'" & CStr(Co.Addr(4)) & "'")
    aFormulaValue.Add CStr(txtPONumber)
    aFormulaValue.Add CStr("'0'")
    aFormulaValue.Add CStr("'1'")
    aFormulaValue.Add CStr("'1'")
    aFormulaValue.Add CStr("'1'")
    aFormulaValue.Add CStr("'1'")
    aFormulaValue.Add CStr("'1'")
    aFormulaValue.Add CStr("'1'")

   If (iUserLogo = 1) Then
      aFormulaValue.Add CStr("'" & CStr(0) & "'")
   Else
      aFormulaValue.Add CStr("'" & CStr(1) & "'")

   End If
   
    aFormulaValue.Add CStr("'" & CStr(GetPreferenceValue("RESALENUMBER", True)) & "'")
    aFormulaName.Add "ShowOurLogo"
    aFormulaValue.Add CStr("'" & CStr(iUserLogo) & "'")
    ' Set Formula values
    cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue


   sSql = "{PohdTable.PONUMBER}=" & Val(txtPONumber) & " "
   
    cCRViewer.SetReportSelectionFormula sSql
    
    cCRViewer.CRViewerSize Me
    cCRViewer.SetDbTableConnection
   
   If optPrn Then
      MarkAsPrinted Val(txtPONumber)
   End If
   
   cCRViewer.OpenCrystalReportObject Me, aFormulaName
   
   cCRViewer.ClearFieldCollection aRptPara
   cCRViewer.ClearFieldCollection aFormulaName
   cCRViewer.ClearFieldCollection aFormulaValue
   
   
   MouseCursor 0
   Exit Sub
   
Ppr01:
   sProcName = "printreport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   Resume Ppr01a
Ppr01a:
   DoModuleErrors Me
   
End Sub


Private Sub MarkAsPrinted(lPonumber As Long)
   Dim sMark As String
   On Error Resume Next
   sMark = "UPDATE PohdTable SET POPRINTED='" & Format(ES_SYSDATE, "mm/dd/yy") & "' " _
           & "WHERE PONUMBER=" & lPonumber & " "
   clsADOCon.ExecuteSql sMark
   
End Sub

Private Sub GetUseLogo()
    Dim RdoLogo As ADODB.Recordset
    Dim bRows As Boolean
    ' Assumed that COMREF is 1 all the time
    sSql = "SELECT ISNULL(COLUSELOGO, 0) as COLUSELOGO FROM ComnTable WHERE COREF = 1"
    bRows = clsADOCon.GetDataSet(sSql, RdoLogo, ES_FORWARD)

    If bRows Then
        With RdoLogo
            iUserLogo = !COLUSELOGO
        End With
        'RdoLogo.Close
        ClearResultSet RdoLogo
    End If
End Sub



