VERSION 5.00
Begin VB.Form PurcPRf12a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Export PO Items to Excel"
   ClientHeight    =   7710
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8325
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7710
   ScaleWidth      =   8325
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdDown 
      Height          =   855
      Left            =   7800
      Picture         =   "PurcPRf12a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   49
      ToolTipText     =   "Move Field Down"
      Top             =   5400
      Width           =   375
   End
   Begin VB.CommandButton cmdUp 
      Height          =   855
      Left            =   7800
      Picture         =   "PurcPRf12a.frx":0093
      Style           =   1  'Graphical
      TabIndex        =   48
      ToolTipText     =   "Move Field Up"
      Top             =   4200
      Width           =   375
   End
   Begin VB.CheckBox cbPrtType 
      Caption         =   "8"
      Height          =   195
      Index           =   7
      Left            =   5760
      TabIndex        =   47
      Top             =   610
      Width           =   375
   End
   Begin VB.CheckBox cbPrtType 
      Caption         =   "7"
      Height          =   195
      Index           =   6
      Left            =   5280
      TabIndex        =   46
      Top             =   610
      Width           =   375
   End
   Begin VB.CheckBox cbPrtType 
      Caption         =   "6"
      Height          =   195
      Index           =   5
      Left            =   4800
      TabIndex        =   45
      Top             =   610
      Width           =   375
   End
   Begin VB.CheckBox cbPrtType 
      Caption         =   "5"
      Height          =   195
      Index           =   4
      Left            =   4320
      TabIndex        =   44
      Top             =   610
      Width           =   375
   End
   Begin VB.CheckBox cbPrtType 
      Caption         =   "4"
      Height          =   195
      Index           =   3
      Left            =   3840
      TabIndex        =   43
      Top             =   610
      Width           =   375
   End
   Begin VB.CheckBox cbPrtType 
      Caption         =   "3"
      Height          =   195
      Index           =   2
      Left            =   3360
      TabIndex        =   42
      Top             =   610
      Width           =   375
   End
   Begin VB.CheckBox cbPrtType 
      Caption         =   "2"
      Height          =   195
      Index           =   1
      Left            =   2880
      TabIndex        =   41
      Top             =   610
      Width           =   375
   End
   Begin VB.CheckBox cbPrtType 
      Caption         =   "1"
      Height          =   195
      Index           =   0
      Left            =   2400
      TabIndex        =   40
      Top             =   610
      Width           =   375
   End
   Begin VB.CheckBox cbPrtTypeALL 
      Caption         =   "ALL"
      Height          =   255
      Left            =   1680
      TabIndex        =   39
      Top             =   600
      Width           =   735
   End
   Begin VB.ComboBox cmbCde 
      Height          =   315
      Left            =   1680
      TabIndex        =   36
      ToolTipText     =   "Enter Product Code to Print (6 Characters)"
      Top             =   960
      Width           =   1935
   End
   Begin VB.ComboBox cmbCls 
      Height          =   315
      Left            =   1680
      TabIndex        =   34
      ToolTipText     =   "Enter Product Class to Print (4 Characters)"
      Top             =   1320
      Width           =   1935
   End
   Begin VB.ComboBox cmbPart 
      Height          =   315
      Left            =   1680
      TabIndex        =   29
      Top             =   240
      Width           =   2655
   End
   Begin VB.CheckBox cbTFasYN 
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   3240
      TabIndex        =   28
      Top             =   7320
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CheckBox cbDescriptiveFieldNames 
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   6840
      TabIndex        =   24
      Top             =   7080
      Width           =   735
   End
   Begin VB.CheckBox cbHeaderRow 
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   3240
      TabIndex        =   20
      Top             =   7080
      Width           =   735
   End
   Begin VB.CommandButton cmdOneToAvail 
      Caption         =   "<----"
      Height          =   375
      Left            =   3600
      TabIndex        =   14
      Top             =   5400
      Width           =   735
   End
   Begin VB.CommandButton cmdAllToAvailable 
      Caption         =   "<===="
      Height          =   375
      Left            =   3600
      TabIndex        =   13
      Top             =   6000
      Width           =   735
   End
   Begin VB.CommandButton cmdAllToExp 
      Caption         =   "====>"
      Height          =   375
      Left            =   3600
      TabIndex        =   12
      Top             =   4200
      Width           =   735
   End
   Begin VB.CommandButton cmdOneToExp 
      Caption         =   "---->"
      Height          =   375
      Left            =   3600
      TabIndex        =   11
      Top             =   4800
      Width           =   735
   End
   Begin VB.ListBox lbExportFields 
      Height          =   2595
      Left            =   4440
      TabIndex        =   10
      Top             =   3960
      Width           =   3255
   End
   Begin VB.ListBox lbAvailableFields 
      Height          =   2595
      Left            =   240
      Sorted          =   -1  'True
      TabIndex        =   9
      Top             =   3960
      Width           =   3255
   End
   Begin VB.ComboBox cmbEndDte 
      Height          =   315
      Left            =   3240
      TabIndex        =   8
      Top             =   2640
      Width           =   1335
   End
   Begin VB.ComboBox cmbStartDte 
      Height          =   315
      Left            =   1080
      TabIndex        =   7
      Top             =   2640
      Width           =   1335
   End
   Begin VB.ComboBox cmbVnd 
      Height          =   315
      Left            =   1080
      TabIndex        =   6
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      Caption         =   "Filter Results by"
      Height          =   735
      Left            =   240
      TabIndex        =   2
      Top             =   1680
      Width           =   5295
      Begin VB.OptionButton optItemRecDate 
         Caption         =   "Item Received Date"
         Height          =   255
         Left            =   3120
         TabIndex        =   5
         Top             =   360
         Width           =   1815
      End
      Begin VB.OptionButton optItemDueDate 
         Caption         =   "Item Due Date"
         Height          =   255
         Left            =   1440
         TabIndex        =   4
         Top             =   360
         Width           =   1455
      End
      Begin VB.OptionButton optPODate 
         Caption         =   "PO Date"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "Export"
      Height          =   375
      Left            =   7200
      TabIndex        =   1
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   7200
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "( Blank For ALL )"
      Height          =   255
      Index           =   16
      Left            =   4440
      TabIndex        =   38
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "( Blank For ALL )"
      Height          =   255
      Index           =   15
      Left            =   4440
      TabIndex        =   37
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Product Code:"
      Height          =   255
      Index           =   13
      Left            =   240
      TabIndex        =   35
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Product Class:"
      Height          =   255
      Index           =   12
      Left            =   240
      TabIndex        =   33
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Part Type(s):"
      Height          =   255
      Index           =   11
      Left            =   240
      TabIndex        =   32
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Part Number:"
      Height          =   255
      Index           =   10
      Left            =   240
      TabIndex        =   31
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "( Blank For ALL )"
      Height          =   255
      Index           =   9
      Left            =   4440
      TabIndex        =   30
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Export True/False Fields as Yes/No"
      Height          =   255
      Index           =   8
      Left            =   360
      TabIndex        =   27
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
      TabIndex        =   26
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
      TabIndex        =   25
      Top             =   3600
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "Use Descriptive Field Names"
      Height          =   255
      Index           =   7
      Left            =   4440
      TabIndex        =   23
      Top             =   7080
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Include Header Row"
      Height          =   255
      Index           =   6
      Left            =   360
      TabIndex        =   22
      Top             =   7080
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Export Options:"
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   21
      Top             =   6720
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "( Blank For ALL )"
      Height          =   255
      Index           =   4
      Left            =   3600
      TabIndex        =   19
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "( Blank For ALL )"
      Height          =   255
      Index           =   3
      Left            =   4920
      TabIndex        =   18
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Vendor:"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   17
      Top             =   3000
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "To:"
      Height          =   255
      Index           =   1
      Left            =   2640
      TabIndex        =   16
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "From:"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   15
      Top             =   2640
      Width           =   735
   End
End
Attribute VB_Name = "PurcPRf12a"
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



Private Sub cbPrtTypeALL_Click()
    EnableDisableFields
End Sub

Private Sub cmbPart_LostFocus()
   cmbPart = CheckLen(cmbPart, 30)
   If cmbPart = "" Then cmbPart = "ALL"
End Sub



Private Sub cmbCls_LostFocus()
    cmbCls = CheckLen(cmbCls, 4)
    If Len(Trim(cmbCls)) = 0 Then cmbCls = "ALL"
End Sub

Private Sub cmbCde_LostFocus()
    cmbCde = CheckLen(cmbCde, 6)
    If Len(Trim(cmbCde)) = 0 Then cmbCde = "ALL"
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
    Dim I As Integer
   
    For I = 0 To lbExportFields.ListCount - 1
      If Not ItemExists(lbAvailableFields, lbExportFields.List(I)) Then lbAvailableFields.AddItem (lbExportFields.List(I))
    Next I
    lbExportFields.Clear
End Sub

Private Sub cmdAllToExp_Click()
    Dim I As Integer
    
    For I = 0 To lbAvailableFields.ListCount - 1
      If Not ItemExists(lbExportFields, lbAvailableFields.List(I)) Then lbExportFields.AddItem (lbAvailableFields.List(I))
    Next I
    lbAvailableFields.Clear
End Sub

Private Sub cmdCan_Click()
  Unload Me
End Sub


Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
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
   Dim RdoPO As ADODB.Recordset
   Dim I As Integer
   Dim sFieldsToExport() As String
   Dim sPartNo As String
   Dim sProdCode As String
   Dim sProdClass As String
   Dim sVendor As String
   Dim bPartTypeSelected As Boolean
   Dim sStartDte, sEndDte As String
   
   On Error GoTo ExportError
   
   If lbExportFields.ListCount <= 0 Then
    MsgBox "Please select at least one field to Export", vbOKOnly
    Exit Sub
   End If
   
   If cbPrtTypeALL.Value <> 1 Then
        For I = 0 To 7
            If cbPrtType(I).Value = 1 Then bPartTypeSelected = True
        Next I
   Else
        bPartTypeSelected = True
   End If
   If Not bPartTypeSelected Then
        MsgBox "You Must Select at Least One Part Type", vbOKOnly
        Exit Sub
   End If
   
   ReDim sFieldsToExport(0 To lbExportFields.ListCount - 1)
    
   cmdExport.Enabled = False

   For I = 0 To lbExportFields.ListCount - 1
    If Len(Trim(lbExportFields.List(I))) > 0 Then sFieldsToExport(I) = ParseFieldName(lbExportFields.List(I))
   Next I
   
   
   
   If Len(cmbPart) = 0 Then cmbPart = "ALL"
   sPartNo = Compress(cmbPart)
   If sPartNo = "ALL" Then sPartNo = ""

   If Len(cmbCde) = 0 Then cmbCde = "ALL"
   sProdCode = Compress(cmbCde)
   If sProdCode = "ALL" Then sProdCode = ""
   
   If Len(cmbCls) = 0 Then cmbCls = "ALL"
   sProdClass = Compress(cmbCls)
   If sProdClass = "ALL" Then sProdClass = ""
   
   If Len(cmbVnd) = 0 Then cmbVnd = "ALL"
   sVendor = Compress(cmbVnd)
   If sVendor = "ALL" Then sVendor = ""
   
   
   sSql = "SELECT  PoitTable.*, PartTable.PACLASS, PartTable.PAPRODCODE, PartTable.PALEVEL, PohdTable.PODATE FROM PohdTable " & _
     " INNER JOIN PoitTable ON PohdTable.PONUMBER = PoiTtable.PINUMBER " & _
     " LEFT OUTER JOIN PartTable ON PartTable.PARTREF = PoitTable.PIPART " & _
     " WHERE "
   
   
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

   ''sSql = "{PoitTable.PIPDATE} "
   If optPODate.Value = True Then
        sSql = sSql & " PohdTable.PODATE "
   ElseIf optItemDueDate.Value = True Then
        sSql = sSql & " PoitTable.PIPDATE "
   Else
        sSql = sSql & " PoitTable.PIADATE "
   End If
   sSql = sSql & " BETWEEN '" & sStartDte & "' AND '" & sEndDte & "' "
   
   
   sSql = sSql & " AND PoitTable.PIVENDOR LIKE '" & sVendor & "%' "
   If sPartNo <> "" Then sSql = sSql & " AND PoitTable.PIPART LIKE '" & sPartNo & "%' "
   If sProdCode <> "" Then sSql = sSql & " AND PartTable.PAPRODCODE LIKE '" & sProdCode & "%' "
   If sProdClass <> "" Then sSql = sSql & " AND PartTable.PACLASS LIKE '" & sProdClass & "%' "
       
   If cbPrtTypeALL.Value <> 1 Then
        sSql = sSql & " AND PartTable.PALEVEL IN ("
        For I = 0 To 7
            sSql = sSql & LTrim(str(I + 1)) & ","
        Next I
        Mid(sSql, Len(sSql), 1) = ")"
   End If
   
   Debug.Print sSql
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPO, ES_STATIC)
   
   If bSqlRows Then SaveAsExcel RdoPO, sFieldsToExport, "", True, (cbHeaderRow.Value = 1), (cbDescriptiveFieldNames.Value = 1), (cbTFasYN.Value = 1) Else MsgBox "No records found. Please try again.", vbOKOnly

   Set RdoPO = Nothing
   cmdExport.Enabled = True
   Exit Sub
   
ExportError:
   MouseCursor 0
   cmdExport.Enabled = True

End Sub


'Private Sub cmdMoveUp_Click()
'    ' only if the first item isn't the current one
'    If lbExportFields.ListIndex > 0 Then
'        ' add a duplicate item up in the listbox
'        lbExportFields.AddItem lbExportFields.Text, lbExportFields.ListIndex - 1
'        ' make it the current item
'        lbExportFields.ListIndex = lbExportFields.ListIndex - 2
'        ' delete the old occurrence of this item
'        lbExportFields.RemoveItem lbExportFields.ListIndex + 2
'    End If
'End Sub
'Private Sub cmdMoveDown_Click()
'    ' only if the last item isn't the current one
'   If lbExportFields.ListIndex <> -1 And lbExportFields.ListIndex < lbExportFields.ListCount - 1 Then
'        ' add a duplicate item down in the listbox
'        lbExportFields.AddItem lbExportFields.Text, lbExportFields.ListIndex + 2
'        ' make it the current item
'        lbExportFields.ListIndex = lbExportFields.ListIndex + 2
'        ' delete the old occurrence of this item
'        lbExportFields.RemoveItem lbExportFields.ListIndex - 2
'    End If
'End Sub


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
    If bOnLoad Then
       FillProductCodes
       FillProductClasses
    End If
    
    bOnLoad = 0
    MouseCursor 0
End Sub


Private Sub Form_Load()
    FormLoad Me
    FillCombo
    lbAvailableFields.Clear
    lbExportFields.Clear
    'FillAvailableFields
    bOnLoad = 1
    FormatControls
    GetOptions
    'If lbExportFields.ListCount = 0 And lbAvailableFields.ListCount = 0 Then FillAvailableFields
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
    Set PurcPRf12a = Nothing
End Sub


Private Sub FillCombo()
   On Error GoTo DiaErr1
   sSql = "Qry_FillVendorsNone"
   cmbVnd.AddItem "ALL"
   LoadComboBox cmbVnd
      
   sSql = "SELECT DISTINCT PIPART,PARTREF,PARTNUM From PoitTable " & _
          "INNER JOIN PartTable ON PoitTable.PIPART = PartTable.PARTREF " & _
          "ORDER BY PIPART"
   LoadComboBox cmbPart, 1
   'If cmbPart.ListCount > 0 Then cmbPart = cmbPart.List(0)
   
    
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub



Private Sub SaveOptions()
    Dim sOptions As String
    Dim I As Integer
    
    
    sOptions = ""
    If optPODate.Value = True Then sOptions = "1" Else If optItemDueDate.Value = True Then sOptions = "2" Else If optItemRecDate.Value = True Then sOptions = "3" Else sOptions = "1"
    sOptions = sOptions & Left(Trim(cmbStartDte) & Space(10), 10)
    sOptions = sOptions & Left(Trim(cmbEndDte) & Space(10), 10)
    sOptions = sOptions & Left(Trim(cmbVnd) & Space(10), 10)
    sOptions = sOptions & cbHeaderRow.Value
    sOptions = sOptions & cbDescriptiveFieldNames.Value
    sOptions = sOptions & cbTFasYN.Value
    sOptions = sOptions & cbPrtTypeALL.Value
    For I = 0 To 7
        sOptions = sOptions & cbPrtType(I).Value
    Next I
    sOptions = sOptions & Left(Trim(cmbCde) & Space(6), 6)
    sOptions = sOptions & Left(Trim(cmbCls) & Space(4), 4)
    sOptions = sOptions & Left(Trim(cmbPart) & Space(30), 30)
    
    SaveLoadListbox lbAvailableFields, "EsiProd", 1
    SaveLoadListbox lbExportFields, "EsiProd", 1
    
    SaveSetting "Esi2000", "EsiProd", "prf12a", Trim(sOptions)
End Sub

Private Sub GetOptions()
   Dim sOptions As String
    Dim I As Integer
    
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiProd", "prf12a", sOptions)
   If Len(sOptions) > 0 Then
        Select Case Val(Left(sOptions, 1))
        Case 1:    optPODate.Value = True
        Case 2:  optItemDueDate.Value = True
        Case 3:  optItemRecDate.Value = True
        End Select
        cmbStartDte = Trim(Mid(sOptions, 2, 10))
        cmbEndDte = Trim(Mid(sOptions, 12, 10))
        cmbVnd = Trim(Mid(sOptions, 22, 10))
        If Len(sOptions) > 31 Then cbHeaderRow.Value = Val(Mid(sOptions, 32, 1)) Else cbHeaderRow.Value = 1
        If Len(sOptions) > 32 Then cbDescriptiveFieldNames.Value = Val(Mid(sOptions, 33, 1)) Else cbDescriptiveFieldNames.Value = 1
        If Len(sOptions) > 33 Then cbTFasYN.Value = Val(Mid(sOptions, 34, 1)) Else cbTFasYN.Value = 1
        If Len(sOptions) > 34 Then cbPrtTypeALL.Value = Val(Mid(sOptions, 35, 1)) Else cbPrtTypeALL.Value = 1
        For I = 0 To 7
            If Len(sOptions) > (35 + I) Then cbPrtType(I).Value = Val(Mid(sOptions, (36 + I), 1)) Else cbPrtType(I + 1).Value = 0
        Next I
        If Len(sOptions) > 43 Then cmbCde = Trim(Mid(sOptions, 44, 6)) Else cmbCde = "ALL"
        If Len(sOptions) > 49 Then cmbCls = Trim(Mid(sOptions, 50, 4)) Else cmbCls = "ALL"
        If Len(sOptions) > 53 Then cmbPart = Trim(Mid(sOptions, 54, 30)) Else cmbPart = "ALL"
    Else
        optPODate.Value = 1
        cmbStartDte = "ALL"
        cmbEndDte = "ALL"
        cmbVnd = "ALL"
        cbHeaderRow.Value = 1
        cbDescriptiveFieldNames.Value = 1
        cbTFasYN.Value = 1
        cbPrtTypeALL.Value = 1
        cmbPart = "ALL"
        For I = 0 To 7
            cbPrtType(I).Value = 0
        Next I
   End If
    
    SaveLoadListbox lbAvailableFields, "EsiProd", 2
    SaveLoadListbox lbExportFields, "EsiProd", 2
    
   
End Sub



Private Sub FillAvailableFields()
    Dim arrTmp() As String
    Dim iLow As Integer
    Dim iHigh As Integer
    
    Dim I As Integer
    Dim i2 As Integer
    Dim sTableName As String
    
    
    Dim bFound As Byte
    
    
    LoadTableColumns "PoitTable", arrTmp()
    iLow = LBound(arrTmp)
    iHigh = UBound(arrTmp)
    iHigh = iHigh + 4
    ReDim Preserve arrTmp(iLow To iHigh)
    
    arrTmp(iHigh - 3) = "PODATE"
    arrTmp(iHigh - 2) = "PACLASS"
    arrTmp(iHigh - 1) = "PAPRODCODE"
    arrTmp(iHigh - 0) = "PALEVEL"
    
    
    bFound = 0
    For I = LBound(arrTmp) To UBound(arrTmp)
        bFound = 0
        If arrTmp(I) = "PODATE" Then
          sTableName = "PohdTable"
        ElseIf arrTmp(I) = "PACLASS" Or arrTmp(I) = "PAPRODCODE" Or arrTmp(I) = "PALEVEL" Then
          sTableName = "PartTable"
        Else
          sTableName = "PoitTable"
        End If

        If UCase(FriendlyFieldName(sTableName, arrTmp(I))) <> "<UNUSED>" Then
            For i2 = 0 To lbAvailableFields.ListCount - 1
                If ParseFieldName(lbAvailableFields.List(i2)) = arrTmp(I) Then
                    bFound = 1
                    lbAvailableFields.List(i2) = FriendlyFieldName(sTableName, arrTmp(I)) & " [" & arrTmp(I) & "]"
                    Exit For
                End If
            Next i2
            If bFound = 0 Then
                For i2 = 0 To lbExportFields.ListCount - 1
                'haven't found it yet, lets look through the selected fields now
                    If ParseFieldName(lbExportFields.List(i2)) = arrTmp(I) Then
                        bFound = 1
                        lbExportFields.List(i2) = FriendlyFieldName(sTableName, arrTmp(I)) & " [" & arrTmp(I) & "]"
                        Exit For
                    End If
                Next i2
            End If
            If bFound = 0 Then lbAvailableFields.AddItem FriendlyFieldName(sTableName, arrTmp(I)) & " [" & arrTmp(I) & "]"
        End If
    Next I


    For I = lbAvailableFields.ListCount - 1 To 0 Step -1
        If ParseFieldName(lbAvailableFields.List(I)) = "PODATE" Then
          sTableName = "PohdTable"
        ElseIf ParseFieldName(lbAvailableFields.List(I)) = "PACLASS" Or ParseFieldName(lbAvailableFields.List(I)) = "PAPRODCODE" Or ParseFieldName(lbAvailableFields.List(I)) = "PALEVEL" Then
          sTableName = "PartTable"
        Else
          sTableName = "PoitTable"
        End If
        If UCase(FriendlyFieldName(sTableName, ParseFieldName(lbAvailableFields.List(I)))) = "<UNUSED>" Then lbAvailableFields.RemoveItem (I)
    Next I
    For I = lbExportFields.ListCount - 1 To 0 Step -1
        If ParseFieldName(lbExportFields.List(I)) = "PODATE" Then
          sTableName = "PohdTable"
        ElseIf ParseFieldName(lbExportFields.List(I)) = "PACLASS" Or ParseFieldName(lbExportFields.List(I)) = "PAPRODCODE" Or ParseFieldName(lbExportFields.List(I)) = "PALEVEL" Then
          sTableName = "PartTable"
        Else
          sTableName = "PoitTable"
        End If
        If UCase(FriendlyFieldName(sTableName, ParseFieldName(lbExportFields.List(I)))) = "<UNUSED>" Then lbExportFields.RemoveItem (I)
    Next I
    

End Sub





Private Sub lbAvailableFields_DblClick()
    If lbAvailableFields.ListIndex = -1 Then Exit Sub
    lbExportFields.AddItem (lbAvailableFields.List(lbAvailableFields.ListIndex))
    lbAvailableFields.RemoveItem (lbAvailableFields.ListIndex)
End Sub


Private Function ItemExists(ByVal lb As ListBox, sItem As String) As Boolean
    Dim I As Integer
    ItemExists = False
    For I = 0 To lb.ListCount - 1
        If lb.List(I) = sItem Then
            ItemExists = True
            Exit For
        End If
    Next I
End Function


Private Sub lbExportFields_DblClick()
    If lbExportFields.ListIndex = -1 Then Exit Sub
    lbAvailableFields.AddItem (lbExportFields.List(lbExportFields.ListIndex))
    lbExportFields.RemoveItem (lbExportFields.ListIndex)

End Sub


Private Sub EnableDisableFields()
    Dim I As Integer
    
    If cbPrtTypeALL.Value = 1 Then
        For I = 0 To 7
            cbPrtType(I).Value = 0
            cbPrtType(I).Enabled = False
        Next I
    Else
        For I = 0 To 7
            cbPrtType(I).Enabled = True
            cbPrtType(I).Value = 1
        Next I
    End If
    



End Sub
