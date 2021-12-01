VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form SaleSLf13a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Export Sales Orders to Excel"
   ClientHeight    =   8310
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8325
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8310
   ScaleWidth      =   8325
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   2040
      TabIndex        =   5
      Top             =   2160
      Width           =   2775
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   1560
      TabIndex        =   40
      Top             =   7920
      Visible         =   0   'False
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdFnd 
      Height          =   375
      Left            =   4920
      Picture         =   "SaleSLf13a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox cmbPrt2 
      Height          =   285
      Left            =   2040
      TabIndex        =   38
      Top             =   2160
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.ComboBox cmbSortDate 
      Height          =   315
      Left            =   2040
      TabIndex        =   2
      Text            =   "cmbSortDate"
      Top             =   1200
      Width           =   2175
   End
   Begin VB.ComboBox cmbSOType 
      Height          =   315
      Left            =   2040
      TabIndex        =   1
      Text            =   "cmbSOType"
      Top             =   720
      Width           =   855
   End
   Begin VB.CheckBox cbCancelled 
      Height          =   255
      Left            =   2040
      TabIndex        =   31
      Top             =   3000
      Width           =   615
   End
   Begin VB.CommandButton cmdDown 
      Height          =   855
      Left            =   7800
      Picture         =   "SaleSLf13a.frx":043A
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "Move Field Down"
      Top             =   5400
      Width           =   375
   End
   Begin VB.CommandButton cmdUp 
      Height          =   855
      Left            =   7800
      Picture         =   "SaleSLf13a.frx":04CD
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "Move Field Up"
      Top             =   4200
      Width           =   375
   End
   Begin VB.ComboBox cmbCust 
      Height          =   315
      Left            =   2040
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
   Begin VB.CheckBox cbTFasYN 
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   3240
      TabIndex        =   26
      Top             =   7320
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CheckBox cbDescriptiveFieldNames 
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   6840
      TabIndex        =   22
      Top             =   7080
      Width           =   735
   End
   Begin VB.CheckBox cbHeaderRow 
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   3240
      TabIndex        =   18
      Top             =   7080
      Width           =   735
   End
   Begin VB.CommandButton cmdOneToAvail 
      Caption         =   "<----"
      Height          =   375
      Left            =   3600
      TabIndex        =   10
      Top             =   5400
      Width           =   735
   End
   Begin VB.CommandButton cmdAllToAvailable 
      Caption         =   "<===="
      Height          =   375
      Left            =   3600
      TabIndex        =   11
      Top             =   6000
      Width           =   735
   End
   Begin VB.CommandButton cmdAllToExp 
      Caption         =   "====>"
      Height          =   375
      Left            =   3600
      TabIndex        =   8
      Top             =   4200
      Width           =   735
   End
   Begin VB.CommandButton cmdOneToExp 
      Caption         =   "---->"
      Height          =   375
      Left            =   3600
      TabIndex        =   9
      Top             =   4800
      Width           =   735
   End
   Begin VB.ListBox lbExportFields 
      Height          =   2595
      Left            =   4440
      TabIndex        =   12
      Top             =   3960
      Width           =   3255
   End
   Begin VB.ListBox lbAvailableFields 
      Height          =   2595
      Left            =   240
      Sorted          =   -1  'True
      TabIndex        =   7
      Top             =   3960
      Width           =   3255
   End
   Begin VB.ComboBox cmbEndDte 
      Height          =   315
      Left            =   4080
      TabIndex        =   4
      Top             =   1680
      Width           =   1335
   End
   Begin VB.ComboBox cmbStartDte 
      Height          =   315
      Left            =   2040
      TabIndex        =   3
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "Export"
      Height          =   375
      Left            =   7200
      TabIndex        =   14
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   7200
      TabIndex        =   13
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblDsc 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2040
      TabIndex        =   39
      Top             =   2520
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "( Blank For ALL )"
      Height          =   255
      Index           =   14
      Left            =   5640
      TabIndex        =   37
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Part Number:"
      Height          =   255
      Index           =   13
      Left            =   240
      TabIndex        =   36
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Date to Sort By:"
      Height          =   255
      Index           =   12
      Left            =   240
      TabIndex        =   35
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "( Blank For ALL )"
      Height          =   255
      Index           =   11
      Left            =   3000
      TabIndex        =   34
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Sales Order Types:"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   33
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Include Cancelled?"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   32
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Customer Nickname:"
      Height          =   255
      Index           =   10
      Left            =   240
      TabIndex        =   28
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "( Blank For ALL )"
      Height          =   255
      Index           =   9
      Left            =   4440
      TabIndex        =   27
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Export True/False Fields as Yes/No"
      Height          =   255
      Index           =   8
      Left            =   360
      TabIndex        =   25
      Top             =   7320
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Fields To Export"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   4440
      TabIndex        =   24
      Top             =   3600
      Width           =   3255
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Available Fields"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   23
      Top             =   3600
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "Use Descriptive Field Names"
      Height          =   255
      Index           =   7
      Left            =   4440
      TabIndex        =   21
      Top             =   7080
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Include Header Row"
      Height          =   255
      Index           =   6
      Left            =   360
      TabIndex        =   20
      Top             =   7080
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Export Options:"
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   19
      Top             =   6720
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "( Blank For ALL )"
      Height          =   255
      Index           =   3
      Left            =   5640
      TabIndex        =   17
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "To:"
      Height          =   255
      Index           =   1
      Left            =   3600
      TabIndex        =   16
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "From:"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   15
      Top             =   1680
      Width           =   735
   End
End
Attribute VB_Name = "SaleSLf13a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bOnLoad As Byte
Dim bCanceled As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd


Private Sub cbHeaderRow_Click()
    If cbHeaderRow.Value = 1 Then
        Label1(7).Enabled = True
        Me.cbDescriptiveFieldNames.Enabled = True
    Else
        Label1(7).Enabled = False
        Me.cbDescriptiveFieldNames.Enabled = False
    End If
End Sub


Private Sub cmbCust_LostFocus()
    If Len(Trim(cmbCust)) = 0 Then cmbCust = "ALL"
End Sub





Private Sub cmbSOType_LostFocus()
    If Len(cmbSOType) = 0 Then cmbSOType.ListIndex = 0
End Sub

Private Sub cmdFnd_Click()
   ViewParts.lblControl = "CMBPRT"
   ViewParts.txtPrt = cmbPrt
   'optVew.Value = vbChecked
   ViewParts.Show
   
End Sub


Private Sub cmbPrt_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF4 Then
      ViewParts.lblControl = "CMBPRT"
      ViewParts.txtPrt = cmbPrt
      'optVew.Value = vbChecked
      ViewParts.Show
   End If
   
End Sub

Private Sub cmbPrt_LostFocus()
   Dim sOldPart As String
   cmbPrt = CheckLen(cmbPrt, 30)
   sOldPart = cmbPrt
   If cmbPrt = "" Then cmbPrt = "ALL"
   If Trim(cmbPrt) <> "ALL" Then
      cmbPrt = CheckLen(cmbPrt, 30)
      cmbPrt = GetCurrentPart(cmbPrt, lblDsc)
      If lblDsc.ForeColor = ES_RED Then lblDsc = ""
      If cmbPrt = "" Then cmbPrt = sOldPart
   Else
      lblDsc = "Range Of Parts Selected."
      cmbPrt = "ALL"
   End If
   
End Sub


Private Sub cmbStartDte_DropDown()
   ShowCalendarEx Me
End Sub


Private Sub cmbStartDte_LostFocus()
   If Len(Trim(cmbStartDte)) = 0 Then cmbStartDte = "ALL"
   If cmbStartDte <> "ALL" Then cmbStartDte = CheckDateEx(cmbStartDte)
End Sub


Private Sub cmbEndDte_DropDown()
   ShowCalendarEx Me
End Sub


Private Sub cmbEndDte_LostFocus()
   If Len(Trim(cmbEndDte)) = 0 Then cmbEndDte = "ALL"
   If cmbEndDte <> "ALL" Then cmbEndDte = CheckDateEx(cmbEndDte)
   
End Sub


Private Sub cmdAllToAvailable_Click()
    Dim i As Integer
   
    For i = 0 To lbExportFields.ListCount - 1
      If Not ItemExists(lbAvailableFields, lbExportFields.List(i)) Then lbAvailableFields.AddItem (lbExportFields.List(i))
    Next i
    lbExportFields.Clear
End Sub

Private Sub cmdAllToExp_Click()
    Dim i As Integer
    
    For i = 0 To lbAvailableFields.ListCount - 1
      If Not ItemExists(lbExportFields, lbAvailableFields.List(i)) Then lbExportFields.AddItem (lbAvailableFields.List(i))
    Next i
    lbAvailableFields.Clear
End Sub

Private Sub cmdCan_Click()
  Unload Me
End Sub


Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   bCanceled = 1
End Sub


Private Sub cmdDown_Click()
   Dim sText As String
   Dim iIndex As Integer
   If lbExportFields.SelCount = 1 Then
        If lbExportFields.ListCount - 1 = lbExportFields.ListIndex Then Exit Sub
        sText = lbExportFields.List(lbExportFields.ListIndex)
        iIndex = lbExportFields.ListIndex
        lbExportFields.RemoveItem lbExportFields.ListIndex
        lbExportFields.AddItem sText, iIndex + 1
        lbExportFields.Selected(iIndex + 1) = True
   End If
End Sub

Private Sub cmdExport_Click()
   Dim RdoSO As ADODB.Recordset
   Dim i As Integer
   Dim sFieldsToExport() As String
   Dim sStartDte, sEndDte As String
   
   On Error GoTo ExportError
   
   If lbExportFields.ListCount <= 0 Then
    MsgBox "Please select at least one field to Export", vbOKOnly
    Exit Sub
   End If
   
   
   ReDim sFieldsToExport(0 To lbExportFields.ListCount - 1)
    
   cmdExport.Enabled = False

   For i = 0 To lbExportFields.ListCount - 1
    If Len(Trim(lbExportFields.List(i))) > 0 Then sFieldsToExport(i) = ParseFieldName(lbExportFields.List(i))
   Next i
   
   
   sSql = "Select * from SohdTable INNER JOIN SoitTable ON SohdTable.SONUMBER = SoitTable.ITSO WHERE "
   
   
   If Len(Trim(cmbStartDte)) = 0 Then cmbStartDte = "ALL"
   If Len(Trim(cmbEndDte)) = 0 Then cmbEndDte = "ALL"
   If Not IsDate(cmbStartDte) Then
      sStartDte = "01/01/1995"
   Else
      sStartDte = Format(cmbStartDte, "mm/dd/yyyy")
   End If
   If Not IsDate(cmbEndDte) Then
      sEndDte = "12/31/2024"
   Else
      sEndDte = Format(cmbEndDte, "mm/dd/yyyy")
   End If
   
   Select Case cmbSortDate.ListIndex
   Case 0: sSql = sSql & "SODATE "
   Case 1: sSql = sSql & "ITBOOKDATE "
   Case 2: sSql = sSql & "ITSCHED "
   Case 3: sSql = sSql & "ITACTUAL "
   Case Else
    sSql = sSql & "SODATE "
   End Select
   sSql = sSql & " BETWEEN '" & sStartDte & "' AND '" & sEndDte & "' "
   If Compress(cmbSOType) <> "ALL" Then sSql = sSql & " AND SohdTable.SOTYPE = '" & Compress(cmbSOType) & "' "
   If Compress(cmbCust) <> "ALL" Then sSql = sSql & " AND SohdTable.SOCUST = '" & Compress(cmbCust) & "' "
   If cbCancelled.Value = 0 Then sSql = sSql & " AND SOCANCELED = 0"
   If Compress(cmbPrt) <> "ALL" Then sSql = sSql & " AND ITPART = '" & Compress(cmbPrt) & "' "
   Debug.Print sSql
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoSO, ES_STATIC)
   
   If bSqlRows Then SaveAsExcel RdoSO, sFieldsToExport, "", True, (cbHeaderRow.Value = 1), (cbDescriptiveFieldNames.Value = 1), (cbTFasYN.Value = 1), ProgressBar1 Else MsgBox "No records found. Please try again.", vbOKOnly

   Set RdoSO = Nothing
   cmdExport.Enabled = True
   Exit Sub
   
ExportError:
   MouseCursor 0
   MsgBox str(Err) & " " & Err.Description
   
   cmdExport.Enabled = True

End Sub



Private Sub cmdOneToAvail_Click()
    If lbExportFields.ListIndex <> -1 Then lbExportFields_DblClick
End Sub

Private Sub cmdOneToExp_Click()
    If lbAvailableFields.ListIndex <> -1 Then lbAvailableFields_DblClick
End Sub


Private Sub cmdUp_Click()
    Dim sText As String
    Dim iIndex As Integer
    
    If lbExportFields.SelCount = 1 Then
        If lbExportFields.ListIndex = 0 Then Exit Sub
        sText = lbExportFields.List(lbExportFields.ListIndex)
        iIndex = lbExportFields.ListIndex
        lbExportFields.RemoveItem lbExportFields.ListIndex
        lbExportFields.AddItem sText, iIndex - 1
        lbExportFields.Selected(iIndex - 1) = True
    End If
End Sub

Private Sub Form_Activate()
   Dim cmbSaved As String
    If bOnLoad Then
     'Fill comboboxes here
     cmbSaved = cmbPrt
     FillPartCombo cmbPrt
     cmbPrt = cmbSaved
    End If
    bOnLoad = 0
    MouseCursor 0
End Sub


Private Sub Form_Load()
    FormLoad Me
    FillCombo
    lbAvailableFields.Clear
    lbExportFields.Clear
    bOnLoad = 1
    FormatControls
    GetOptions
    FillAvailableFields
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveOptions
End Sub

Private Sub Form_Resize()
    Refresh
End Sub


Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FormUnload
    Set SaleSLf13a = Nothing
End Sub


Private Sub FillCombo()
   Dim i As Integer
   On Error GoTo DiaErr1

   sSql = "SELECT DISTINCT SOCUST FROM SohdTable ORDER BY SOCUST"
   LoadComboBox cmbCust, -1


    cmbSOType.Clear
    cmbSOType.AddItem "ALL"
    For i = 65 To 90
        cmbSOType.AddItem Chr(i)
    Next i
    
    cmbSortDate.Clear
    cmbSortDate.AddItem "Sales Order Date"
    cmbSortDate.AddItem "Booking Date"
    cmbSortDate.AddItem "Ship Date"
    cmbSortDate.AddItem "Delivery Date"


   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub



Private Sub SaveOptions()
    Dim sOptions As String
    Dim i As Integer
    
    
    sOptions = ""
    sOptions = sOptions & Left(Trim(cmbStartDte) & Space(10), 10)
    sOptions = sOptions & Left(Trim(cmbEndDte) & Space(10), 10)
    sOptions = sOptions & cbHeaderRow.Value
    sOptions = sOptions & cbDescriptiveFieldNames.Value
    sOptions = sOptions & cbTFasYN.Value
    sOptions = sOptions & cbCancelled.Value
    sOptions = sOptions & Left(cmbCust & Space(30), 30)
    sOptions = sOptions & Right("00" & LTrim(str(cmbSOType.ListIndex)), 2)
    sOptions = sOptions & Right("0" & LTrim(str(cmbSortDate.ListIndex)), 1)
    sOptions = sOptions & Left(cmbPrt & Space(30), 30)
    
    SaveLoadListbox lbAvailableFields, "EsiSalesf13a", 1
    SaveLoadListbox lbExportFields, "EsiSalesf13a", 1
    
    SaveSetting "Esi2000", "EsiSales", "slf13a", Trim(sOptions)
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   Dim i As Integer
    
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiSales", "slf13a", sOptions)
   If Len(sOptions) > 0 Then
        cmbStartDte = Trim(Mid(sOptions, 1, 10))
        cmbEndDte = Trim(Mid(sOptions, 11, 10))
        If Len(sOptions) > 20 Then cbHeaderRow.Value = Val(Mid(sOptions, 21, 1)) Else cbHeaderRow.Value = 1
        If Len(sOptions) > 21 Then cbDescriptiveFieldNames.Value = Val(Mid(sOptions, 22, 1)) Else cbDescriptiveFieldNames.Value = 1
        If Len(sOptions) > 22 Then cbTFasYN.Value = Val(Mid(sOptions, 23, 1)) Else cbTFasYN.Value = 1
        If Len(sOptions) > 23 Then cbCancelled.Value = Val(Mid(sOptions, 24, 1)) Else cbCancelled.Value = 0
        If Len(sOptions) > 24 Then cmbCust = Trim(Mid(sOptions, 25, 30)) Else cmbCust = "ALL"
        If Len(sOptions) > 55 Then cmbSOType.ListIndex = Val(Mid(sOptions, 55, 2)) Else cmbSOType.ListIndex = 0
        If Len(sOptions) > 56 Then cmbSortDate.ListIndex = Val(Mid(sOptions, 57, 1)) Else cmbSortDate.ListIndex = 0
        If Len(sOptions) > 57 Then cmbPrt = Trim(Mid(sOptions, 58, 30)) Else cmbPrt = "ALL"
    Else
        cmbStartDte = "ALL"
        cmbEndDte = "ALL"
        cbHeaderRow.Value = 1
        cbDescriptiveFieldNames.Value = 1
        cbTFasYN.Value = 1
        cbCancelled.Value = 0
        cmbCust = "ALL"
        cmbSOType.ListIndex = 0
        cmbSortDate.ListIndex = 0
        cmbPrt = "ALL"
   End If
    
    SaveLoadListbox lbAvailableFields, "EsiSalesf13a", 2
    SaveLoadListbox lbExportFields, "EsiSalesf13a", 2
    
   
End Sub



Private Sub FillAvailableFields()
    Dim arrTmp() As String
    Dim iLow As Integer
    Dim iHigh As Integer
    
    Dim i As Integer
    Dim i2 As Integer
    Dim sTableName As String
    
    
    Dim bFound As Byte
    
    
    LoadTableColumns "SoitTable", arrTmp()
    iLow = LBound(arrTmp)
    iHigh = UBound(arrTmp)
    iHigh = iHigh + 24
    ReDim Preserve arrTmp(iLow To iHigh)
        
    arrTmp(iHigh - 23) = "SONUMBER"
    arrTmp(iHigh - 22) = "SOTYPE"
    arrTmp(iHigh - 21) = "SOCUST"
    arrTmp(iHigh - 20) = "SODATE"
    arrTmp(iHigh - 19) = "SOSALESMAN"
    arrTmp(iHigh - 18) = "SOREP"
    arrTmp(iHigh - 17) = "SOPO"
    arrTmp(iHigh - 16) = "SOSTNAME"
    arrTmp(iHigh - 15) = "SOSTADR"
    arrTmp(iHigh - 14) = "SOCCONTACT"
    arrTmp(iHigh - 13) = "SOCPHONE"
    arrTmp(iHigh - 12) = "SOCEXT"
    arrTmp(iHigh - 11) = "SOJOBNO"
    arrTmp(iHigh - 10) = "SODIVISION"
    arrTmp(iHigh - 9) = "SOREGION"
    arrTmp(iHigh - 8) = "SOBUSUNIT"
    arrTmp(iHigh - 7) = "SODELDATE"
    arrTmp(iHigh - 6) = "SOCREATED"
    arrTmp(iHigh - 5) = "SOREVISED"
    arrTmp(iHigh - 4) = "SOREMARKS"
    arrTmp(iHigh - 3) = "SOSHIPDATE"
    arrTmp(iHigh - 2) = "SOFREIGHTDAYS"
    arrTmp(iHigh - 1) = "SOCANDATE"
    arrTmp(iHigh - 0) = "SOCANCELED"
    
    
    
    bFound = 0
    For i = LBound(arrTmp) To UBound(arrTmp)
        bFound = 0
        If arrTmp(i) = "SONUMBER" Or arrTmp(i) = "SOTYPE" Or arrTmp(i) = "SOCUST" Or arrTmp(i) = "SODATE" Or arrTmp(i) = "SOSALESMAN" Or arrTmp(i) = "SOREP" _
           Or arrTmp(i) = "SOPO" Or arrTmp(i) = "SOSTNAME" Or arrTmp(i) = "SOSTADR" Or arrTmp(i) = "SOCCONTACT" Or arrTmp(i) = "SOCPHONE" Or arrTmp(i) = "SOCEXT" _
           Or arrTmp(i) = "SOJOBNO" Or arrTmp(i) = "SODIVISION" Or arrTmp(i) = "SOREGION" Or arrTmp(i) = "SOBUSUNIT" Or arrTmp(i) = "SODELDATE" Or arrTmp(i) = "SOCREATED" _
           Or arrTmp(i) = "SOREVISED" Or arrTmp(i) = "SOREMARKS" Or arrTmp(i) = "SOSHIPDATE" Or arrTmp(i) = "SOFREIGHTDAYS" Or arrTmp(i) = "SOCANDATE" Or arrTmp(i) = "SOCANCELED" Then
          sTableName = "SohdTable"
        Else
          sTableName = "SoitTable"
        End If

        If UCase(FriendlyFieldName(sTableName, arrTmp(i))) <> "<UNUSED>" Then
            For i2 = 0 To lbAvailableFields.ListCount - 1
                If ParseFieldName(lbAvailableFields.List(i2)) = arrTmp(i) Then
                    bFound = 1
                    lbAvailableFields.List(i2) = FriendlyFieldName(sTableName, arrTmp(i)) & " [" & arrTmp(i) & "]"
                    Exit For
                End If
            Next i2
            If bFound = 0 Then
                For i2 = 0 To lbExportFields.ListCount - 1
                'haven't found it yet, lets look through the selected fields now
                    If ParseFieldName(lbExportFields.List(i2)) = arrTmp(i) Then
                        bFound = 1
                        lbExportFields.List(i2) = FriendlyFieldName(sTableName, arrTmp(i)) & " [" & arrTmp(i) & "]"
                        Exit For
                    End If
                Next i2
            End If
            If bFound = 0 Then lbAvailableFields.AddItem FriendlyFieldName(sTableName, arrTmp(i)) & " [" & arrTmp(i) & "]"
        End If
    Next i


    For i = lbAvailableFields.ListCount - 1 To 0 Step -1
        If ParseFieldName(lbAvailableFields.List(i)) = "SONUMBER" Or ParseFieldName(lbAvailableFields.List(i)) = "SOTYPE" Or ParseFieldName(lbAvailableFields.List(i)) = "SOCUST" Or ParseFieldName(lbAvailableFields.List(i)) = "SODATE" Or ParseFieldName(lbAvailableFields.List(i)) = "SOSALESMAN" Or ParseFieldName(lbAvailableFields.List(i)) = "SOREP" _
           Or ParseFieldName(lbAvailableFields.List(i)) = "SOPO" Or ParseFieldName(lbAvailableFields.List(i)) = "SOSTNAME" Or ParseFieldName(lbAvailableFields.List(i)) = "SOSTADR" Or ParseFieldName(lbAvailableFields.List(i)) = "SOCCONTACT" Or ParseFieldName(lbAvailableFields.List(i)) = "SOCPHONE" Or ParseFieldName(lbAvailableFields.List(i)) = "SOCEXT" _
           Or ParseFieldName(lbAvailableFields.List(i)) = "SOJOBNO" Or ParseFieldName(lbAvailableFields.List(i)) = "SODIVISION" Or ParseFieldName(lbAvailableFields.List(i)) = "SOREGION" Or ParseFieldName(lbAvailableFields.List(i)) = "SOBUSUNIT" Or ParseFieldName(lbAvailableFields.List(i)) = "SODELDATE" Or ParseFieldName(lbAvailableFields.List(i)) = "SOCREATED" _
           Or ParseFieldName(lbAvailableFields.List(i)) = "SOREVISED" Or ParseFieldName(lbAvailableFields.List(i)) = "SOREMARKS" Or ParseFieldName(lbAvailableFields.List(i)) = "SOSHIPDATE" Or ParseFieldName(lbAvailableFields.List(i)) = "SOFREIGHTDAYS" Or ParseFieldName(lbAvailableFields.List(i)) = "SOCANDATE" Or ParseFieldName(lbAvailableFields.List(i)) = "SOCANCELED" Then
          sTableName = "SohdTable"
        Else
          sTableName = "SoitTable"
        End If
        If UCase(FriendlyFieldName(sTableName, ParseFieldName(lbAvailableFields.List(i)))) = "<UNUSED>" Then lbAvailableFields.RemoveItem (i)
    Next i
    For i = lbExportFields.ListCount - 1 To 0 Step -1
        
        If ParseFieldName(lbExportFields.List(i)) = "SONUMBER" Or ParseFieldName(lbExportFields.List(i)) = "SOTYPE" Or ParseFieldName(lbExportFields.List(i)) = "SOCUST" Or ParseFieldName(lbExportFields.List(i)) = "SODATE" Or ParseFieldName(lbExportFields.List(i)) = "SOSALESMAN" Or ParseFieldName(lbExportFields.List(i)) = "SOREP" _
           Or ParseFieldName(lbExportFields.List(i)) = "SOPO" Or ParseFieldName(lbExportFields.List(i)) = "SOSTNAME" Or ParseFieldName(lbExportFields.List(i)) = "SOSTADR" Or ParseFieldName(lbExportFields.List(i)) = "SOCCONTACT" Or ParseFieldName(lbExportFields.List(i)) = "SOCPHONE" Or ParseFieldName(lbExportFields.List(i)) = "SOCEXT" _
           Or ParseFieldName(lbExportFields.List(i)) = "SOJOBNO" Or ParseFieldName(lbExportFields.List(i)) = "SODIVISION" Or ParseFieldName(lbExportFields.List(i)) = "SOREGION" Or ParseFieldName(lbExportFields.List(i)) = "SOBUSUNIT" Or ParseFieldName(lbExportFields.List(i)) = "SODELDATE" Or ParseFieldName(lbExportFields.List(i)) = "SOCREATED" _
           Or ParseFieldName(lbExportFields.List(i)) = "SOREVISED" Or ParseFieldName(lbExportFields.List(i)) = "SOREMARKS" Or ParseFieldName(lbExportFields.List(i)) = "SOSHIPDATE" Or ParseFieldName(lbExportFields.List(i)) = "SOFREIGHTDAYS" Or ParseFieldName(lbExportFields.List(i)) = "SOCANDATE" Or ParseFieldName(lbExportFields.List(i)) = "SOCANCELED" Then
          sTableName = "SohdTable"
        Else
          sTableName = "SoitTable"
        End If
        If UCase(FriendlyFieldName(sTableName, ParseFieldName(lbExportFields.List(i)))) = "<UNUSED>" Then lbExportFields.RemoveItem (i)
    Next i

End Sub



Private Sub lbAvailableFields_DblClick()
    If lbAvailableFields.ListIndex = -1 Then Exit Sub
    lbExportFields.AddItem (lbAvailableFields.List(lbAvailableFields.ListIndex))
    lbAvailableFields.RemoveItem (lbAvailableFields.ListIndex)
End Sub


Private Function ItemExists(ByVal lb As ListBox, sItem As String) As Boolean
    Dim i As Integer
    ItemExists = False
    For i = 0 To lb.ListCount - 1
        If lb.List(i) = sItem Then
            ItemExists = True
            Exit For
        End If
    Next i
End Function


Private Sub lbExportFields_DblClick()
    If lbExportFields.ListIndex = -1 Then Exit Sub
    lbAvailableFields.AddItem (lbExportFields.List(lbExportFields.ListIndex))
    lbExportFields.RemoveItem (lbExportFields.ListIndex)

End Sub



