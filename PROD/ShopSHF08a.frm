VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form ShopSHF08a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Export MO's to Excel"
   ClientHeight    =   8295
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8460
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   8460
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   2040
      TabIndex        =   44
      Top             =   7920
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdDown 
      Height          =   855
      Left            =   8040
      Picture         =   "ShopSHF08a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   43
      ToolTipText     =   "Move Field Down"
      Top             =   5280
      Width           =   375
   End
   Begin VB.CommandButton cmdUp 
      Height          =   855
      Left            =   8040
      Picture         =   "ShopSHF08a.frx":0093
      Style           =   1  'Graphical
      TabIndex        =   42
      ToolTipText     =   "Move Field Up"
      Top             =   4080
      Width           =   375
   End
   Begin VB.ComboBox cmbCls 
      Height          =   315
      Left            =   1200
      TabIndex        =   10
      ToolTipText     =   "Enter Product Class or Blank for ALL"
      Top             =   1800
      Width           =   1455
   End
   Begin VB.ComboBox cmbCde 
      Height          =   315
      Left            =   1200
      TabIndex        =   9
      ToolTipText     =   "Enter Product Code or Blank for ALL"
      Top             =   1320
      Width           =   1455
   End
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   1200
      TabIndex        =   8
      ToolTipText     =   "Enter Partial Part Number, Entire Part Number, or Blank for ALL"
      Top             =   840
      Width           =   3135
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   1440
      TabIndex        =   34
      Top             =   2160
      Width           =   6255
      Begin VB.OptionButton optDate 
         Caption         =   "Scheduled Completion Date"
         Height          =   195
         Index           =   1
         Left            =   3000
         TabIndex        =   12
         Top             =   240
         Width           =   2535
      End
      Begin VB.OptionButton optDate 
         Caption         =   "Scheduled Start Date"
         Height          =   195
         Index           =   0
         Left            =   840
         TabIndex        =   11
         Top             =   240
         Width           =   1935
      End
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   7800
      Top             =   6600
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   8295
      FormDesignWidth =   8460
   End
   Begin VB.CommandButton cmbExport 
      Caption         =   "Export"
      Height          =   375
      Left            =   7200
      TabIndex        =   33
      Top             =   480
      Width           =   1095
   End
   Begin VB.ListBox lbAvailableFields 
      Height          =   2595
      Left            =   480
      Sorted          =   -1  'True
      TabIndex        =   15
      Top             =   3840
      Width           =   3255
   End
   Begin VB.ListBox lbExportFields 
      Height          =   2595
      Left            =   4680
      TabIndex        =   20
      Top             =   3840
      Width           =   3255
   End
   Begin VB.CommandButton cmdOneToExp 
      Caption         =   "---->"
      Height          =   375
      Left            =   3840
      TabIndex        =   17
      Top             =   4680
      Width           =   735
   End
   Begin VB.CommandButton cmdAllToExp 
      Caption         =   "====>"
      Height          =   375
      Left            =   3840
      TabIndex        =   16
      Top             =   4080
      Width           =   735
   End
   Begin VB.CommandButton cmdAllToAvailable 
      Caption         =   "<===="
      Height          =   375
      Left            =   3840
      TabIndex        =   19
      Top             =   5880
      Width           =   735
   End
   Begin VB.CommandButton cmdOneToAvail 
      Caption         =   "<----"
      Height          =   375
      Left            =   3840
      TabIndex        =   18
      Top             =   5280
      Width           =   735
   End
   Begin VB.CheckBox cbHeaderRow 
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   3480
      TabIndex        =   21
      Top             =   6960
      Width           =   735
   End
   Begin VB.CheckBox cbDescriptiveFieldNames 
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   7080
      TabIndex        =   22
      Top             =   6960
      Width           =   735
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   7200
      TabIndex        =   27
      Top             =   0
      Width           =   1095
   End
   Begin VB.ComboBox txtEnd 
      Height          =   315
      Left            =   5040
      TabIndex        =   14
      Top             =   2880
      Width           =   1455
   End
   Begin VB.ComboBox txtBeg 
      Height          =   315
      Left            =   2280
      TabIndex        =   13
      Top             =   2880
      Width           =   1455
   End
   Begin VB.CheckBox optSta 
      Caption         =   "CA"
      Height          =   255
      Index           =   7
      Left            =   6120
      TabIndex        =   7
      Top             =   360
      Width           =   615
   End
   Begin VB.CheckBox optSta 
      Caption         =   "CL"
      Height          =   255
      Index           =   6
      Left            =   5520
      TabIndex        =   6
      Top             =   360
      Width           =   615
   End
   Begin VB.CheckBox optSta 
      Caption         =   "CO"
      Height          =   255
      Index           =   5
      Left            =   4800
      TabIndex        =   5
      Top             =   360
      Width           =   615
   End
   Begin VB.CheckBox optSta 
      Caption         =   "PC"
      Height          =   255
      Index           =   4
      Left            =   4080
      TabIndex        =   4
      Top             =   360
      Width           =   615
   End
   Begin VB.CheckBox optSta 
      Caption         =   "PP"
      Height          =   255
      Index           =   3
      Left            =   3360
      TabIndex        =   3
      Top             =   360
      Width           =   615
   End
   Begin VB.CheckBox optSta 
      Caption         =   "PL"
      Height          =   255
      Index           =   2
      Left            =   2640
      TabIndex        =   2
      Top             =   360
      Width           =   615
   End
   Begin VB.CheckBox optSta 
      Caption         =   "RL"
      Height          =   255
      Index           =   1
      Left            =   1920
      TabIndex        =   1
      Top             =   360
      Width           =   615
   End
   Begin VB.CheckBox optSta 
      Caption         =   "SC"
      Height          =   255
      Index           =   0
      Left            =   1200
      TabIndex        =   0
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "(Blank For All)"
      Height          =   255
      Index           =   13
      Left            =   4560
      TabIndex        =   41
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "(Blank For All)"
      Height          =   255
      Index           =   12
      Left            =   4560
      TabIndex        =   40
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "(Partial Number or Blank For All)"
      Height          =   255
      Index           =   11
      Left            =   4560
      TabIndex        =   39
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Product Class"
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   38
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Product Code"
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   37
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Part Number"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   36
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Select MO's by"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   35
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Export Options:"
      Height          =   255
      Index           =   5
      Left            =   480
      TabIndex        =   32
      Top             =   6600
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Include Header Row"
      Height          =   255
      Index           =   6
      Left            =   600
      TabIndex        =   31
      Top             =   6960
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Use Descriptive Field Names"
      Height          =   255
      Index           =   7
      Left            =   4680
      TabIndex        =   30
      Top             =   6960
      Width           =   2295
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
      Left            =   480
      TabIndex        =   29
      Top             =   3480
      Width           =   3255
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
      Left            =   4680
      TabIndex        =   28
      Top             =   3480
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "(Blank For All)"
      Height          =   255
      Index           =   3
      Left            =   6840
      TabIndex        =   26
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Through"
      Height          =   255
      Index           =   2
      Left            =   3960
      TabIndex        =   25
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Manufacturing Orders From"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   24
      Top             =   2880
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Run Status:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   23
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "ShopSHF08a"
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

Private txtKeyPress(2) As New EsiKeyBd
Private txtGotFocus(2) As New EsiKeyBd




Private Sub cbHeaderRow_Click()
    If cbHeaderRow.Value = 1 Then
        Label1(7).Enabled = True
        Me.cbDescriptiveFieldNames.Enabled = True
        
    Else
        Label1(7).Enabled = False
        Me.cbDescriptiveFieldNames.Enabled = False
    End If
    
End Sub


Private Sub cmbCde_LostFocus()
    If Len(cmbCde) = 0 Then cmbCde = "ALL"
End Sub



Private Sub cmbCls_LostFocus()
    If Len(cmbCls) = 0 Then cmbCls = "ALL"
End Sub

Private Sub cmbExport_Click()
   Dim RdoPO As ADODB.Recordset
   Dim i As Integer
   Dim sFieldsToExport() As String
   
   Dim sStartDte, sEndDte As String
   
   On Error GoTo ExportError
   
   If lbExportFields.ListCount <= 0 Then
    MsgBox "Please select at least one field to Export", vbOKOnly
    Exit Sub
   End If
      
   ReDim sFieldsToExport(0 To lbExportFields.ListCount - 1)
   cmbExport.Enabled = False

   For i = 0 To lbExportFields.ListCount - 1
    If Len(Trim(lbExportFields.List(i))) > 0 Then sFieldsToExport(i) = ParseFieldName(lbExportFields.List(i))
   Next i
   
   sSql = "SELECT PartTable.PARTNUM, PartTable.PADESC, PartTable.PACLASS, PartTable.PAPRODCODE, RunsTable.* From RunsTable  "
   sSql = sSql & " INNER JOIN PartTable on PartTable.PARTREF=RunsTable.RUNREF "
   sSql = sSql & " WHERE "
   
   If Len(Trim(txtBeg)) = 0 Then txtBeg = "ALL"
   If Len(Trim(txtEnd)) = 0 Then txtEnd = "ALL"
   If Not IsDate(txtBeg) Then
      sStartDte = "01/01/1995"
   Else
      sStartDte = Format(txtBeg, "mm/dd/yyyy")
   End If
   If Not IsDate(txtEnd) Then
      sEndDte = "12/31/2024"
   Else
      sEndDte = Format(txtEnd, "mm/dd/yyyy")
   End If

   If optDate(0).Value = True Then sSql = sSql & " RUNSCHED " Else sSql = sSql + " RUNCOMPLETE "
   sSql = sSql & " BETWEEN '" & sStartDte & "' AND '" & sEndDte & "' "
   
   If optSta(0).Value = vbUnchecked Then sSql = sSql & " AND RUNSTATUS<>'SC' "
   If optSta(1).Value = vbUnchecked Then sSql = sSql & " AND RUNSTATUS<>'RL' "
   If optSta(2).Value = vbUnchecked Then sSql = sSql & " AND RUNSTATUS<>'PL' "
   If optSta(3).Value = vbUnchecked Then sSql = sSql & " AND RUNSTATUS<>'PP' "
   If optSta(4).Value = vbUnchecked Then sSql = sSql & " AND RUNSTATUS<>'PC' "
   If optSta(5).Value = vbUnchecked Then sSql = sSql & " AND RUNSTATUS<>'CO' "
   If optSta(6).Value = vbUnchecked Then sSql = sSql & " AND RUNSTATUS<>'CL' "
   If optSta(7).Value = vbUnchecked Then sSql = sSql & " AND RUNSTATUS<>'CA' "
      
   If Len(cmbCls) > 0 And Compress(cmbCls) <> "ALL" Then sSql = sSql & " AND PartTable.PACLASS = '" & cmbCls & "' "
   If Len(cmbCde) > 0 And Compress(cmbCde) <> "ALL" Then sSql = sSql & " AND PartTable.PAPRODCODE = '" & cmbCde & "' "
   If Len(cmbPrt) > 0 And Compress(cmbPrt) <> "ALL" Then sSql = sSql & " AND PartTable.PARTREF LIKE '" & Compress(cmbPrt) & "%' "
   
   Debug.Print sSql
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPO, ES_STATIC)
   
   If bSqlRows Then SaveAsExcel RdoPO, sFieldsToExport, "", True, (cbHeaderRow.Value = 1), (cbDescriptiveFieldNames.Value = 1), False, ProgressBar1 Else MsgBox "No records found. Please try again.", vbOKOnly

   Set RdoPO = Nothing
   cmbExport.Enabled = True
   Exit Sub
   
ExportError:
   MouseCursor 0
   cmbExport.Enabled = True
   MsgBox Err.Description
   

End Sub


Private Sub cmbPrt_LostFocus()
    If Len(cmbPrt) = 0 Then cmbPrt = "ALL"
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

Private Sub cmdClose_Click()
   Unload Me
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
   MDISect.lblBotPanel = Caption
   MouseCursor vbHourglass
   If bOnLoad Then
      bOnLoad = 0
      GetOptions
      FillProductCodes
      FillProductClasses
      FillAllRuns cmbPrt
      'If lbExportFields.ListCount = 0 And lbAvailableFields.ListCount = 0 Then FillAvailableFields
      FillAvailableFields
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   bOnLoad = 1
   Show
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveOptions
End Sub

Private Sub Form_Resize()
    Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set ShopSHF08a = Nothing
End Sub


Private Sub FormatControls()
   txtBeg = Format(ES_SYSDATE, "mm/01/yyyy")
   txtEnd = Format(ES_SYSDATE, "mm/dd/yyyy")
   
End Sub

Private Sub SaveOptions()
   Dim iList As Integer
   Dim sOptions As String
   
   'Save by Menu Option
   For iList = 0 To 6
      sOptions = sOptions & Trim(str(optSta(iList).Value))
   Next
   sOptions = sOptions & Left(txtBeg & Space(10), 10)
   sOptions = sOptions & Left(txtEnd & Space(10), 10)
   sOptions = sOptions & cbHeaderRow.Value
   sOptions = sOptions & cbDescriptiveFieldNames.Value
   
   For iList = 0 To 1
    If optDate(iList).Value = True Then sOptions = sOptions & "1" Else sOptions = sOptions & "0"
   Next iList
   
   sOptions = sOptions & Left(cmbPrt & Space(30), 30)
   sOptions = sOptions & Left(cmbCde & Space(6), 6)
   sOptions = sOptions & Left(cmbCls & Space(4), 4)
   
   
   SaveSetting "Esi2000", "EsiProd", "shopshf08a", Trim(sOptions)
   SaveLoadListbox lbAvailableFields, "ShopSHf08a", 1
   SaveLoadListbox lbExportFields, "ShopSHf08a", 1
  
End Sub

Private Sub GetOptions()
   Dim iList As Integer
   Dim sOptions As String
   
   On Error Resume Next
   'Get By Menu Option
   sOptions = GetSetting("Esi2000", "EsiProd", "shopshf08a", sOptions)
   If Len(sOptions) > 0 Then
      For iList = 0 To 6
         optSta(iList) = Mid$(sOptions, iList + 1, 1)
      Next
      txtBeg = Trim(Mid(sOptions, 8, 10))
      txtEnd = Trim(Mid(sOptions, 18, 10))
      If Len(sOptions) > 27 Then cbHeaderRow.Value = Val(Mid(sOptions, 28, 1)) Else cbHeaderRow.Value = 1
      If Len(sOptions) > 28 Then cbDescriptiveFieldNames.Value = Val(Mid(sOptions, 29, 1)) Else cbDescriptiveFieldNames.Value = 1
      For iList = 0 To 1
       If Mid(sOptions, iList + 30, 1) = "1" Then optDate(iList).Value = True Else optDate(iList).Value = False
      Next iList
      cmbPrt = Trim(Mid(sOptions, 32, 30))
      cmbCde = Trim(Mid(sOptions, 62, 6))
      cmbCls = Trim(Mid(sOptions, 68, 4))

   Else
      For iList = 0 To 6
         optSta(iList).Value = vbChecked
      Next
      cbHeaderRow.Value = 1
      cbDescriptiveFieldNames.Value = 1
      optDate(0).Value = True
      cmbPrt = "ALL"
      cmbCde = "ALL"
      cmbCls = "ALL"
   End If
   SaveLoadListbox lbAvailableFields, "ShopSHf08a", 2
   SaveLoadListbox lbExportFields, "ShopSHf08a", 2
   
End Sub





Private Sub txtBeg_DropDown()
   ShowCalendarEx Me
End Sub

Private Sub txtBeg_LostFocus()
   If Len(Trim(txtBeg)) = 0 Then txtBeg = "ALL"
   If txtBeg <> "ALL" Then txtBeg = CheckDateEx(txtBeg)
End Sub


Private Sub txtend_DropDown()
   ShowCalendarEx Me
End Sub

Private Sub txtEnd_LostFocus()
   If Len(Trim(txtEnd)) = 0 Then txtEnd = "ALL"
   If txtEnd <> "ALL" Then txtEnd = CheckDateEx(txtEnd)
End Sub


Private Sub FillAvailableFields()
    Dim arrTmp() As String
    Dim iLow As Integer
    Dim iHigh As Integer
    
    Dim i As Integer
    Dim i2 As Integer
    Dim sTableName As String
    
    
    Dim bFound As Byte
    
    
    LoadTableColumns "RunsTable", arrTmp()
    iLow = LBound(arrTmp)
    iHigh = UBound(arrTmp)
    iHigh = iHigh + 4
    ReDim Preserve arrTmp(iLow To iHigh)
    
    'arrTmp(iHigh - 2) = FriendlyFieldName("PartTable", "PARTNUM") & " [PARTNUM]"
    'arrTmp(iHigh - 1) = FriendlyFieldName("PartTable", "PACLASS") & " [PACLASS]"
    'arrTmp(iHigh - 0) = FriendlyFieldName("PartTable", "PAPRODCODE") & " [PAPRODCODE]"
    arrTmp(iHigh - 3) = "PARTNUM"
    arrTmp(iHigh - 2) = "PADESC"
    arrTmp(iHigh - 1) = "PACLASS"
    arrTmp(iHigh - 0) = "PAPRODCODE"
    
    
    bFound = 0
    For i = LBound(arrTmp) To UBound(arrTmp)
        bFound = 0
        If arrTmp(i) = "PARTNUM" Or arrTmp(i) = "PACLASS" Or arrTmp(i) = "PAPRODCODE" Or arrTmp(i) = "PADESC" Then sTableName = "PARTTABLE" Else sTableName = "RUNSTABLE"
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
        If ParseFieldName(lbAvailableFields.List(i)) = "PARTNUM" Or ParseFieldName(lbAvailableFields.List(i)) = "PACLASS" Or ParseFieldName(lbAvailableFields.List(i)) = "PAPRODCODE" Or ParseFieldName(lbAvailableFields.List(i)) = "PADESC" Then sTableName = "PARTTABLE" Else sTableName = "RUNSTABLE"
        If UCase(FriendlyFieldName(sTableName, ParseFieldName(lbAvailableFields.List(i)))) = "<UNUSED>" Then lbAvailableFields.RemoveItem (i)
    Next i
    For i = lbExportFields.ListCount - 1 To 0 Step -1
        If ParseFieldName(lbExportFields.List(i)) = "PARTNUM" Or ParseFieldName(lbExportFields.List(i)) = "PACLASS" Or ParseFieldName(lbExportFields.List(i)) = "PAPRODCODE" Or ParseFieldName(lbExportFields.List(i)) = "PADESC" Then sTableName = "PARTTABLE" Else sTableName = "RUNSTABLE"
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
