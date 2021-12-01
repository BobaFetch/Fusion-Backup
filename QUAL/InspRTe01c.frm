VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form InspRTe01c 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Inspection Tag Items"
   ClientHeight    =   9495
   ClientLeft      =   2100
   ClientTop       =   1635
   ClientWidth     =   7995
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   Icon            =   "InspRTe01c.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9495
   ScaleMode       =   0  'User
   ScaleWidth      =   7995
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtMatType 
      Height          =   825
      Left            =   3960
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   43
      Tag             =   "9"
      Top             =   5640
      Width           =   3795
   End
   Begin VB.TextBox txtRev 
      Height          =   825
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   42
      Tag             =   "9"
      Top             =   5670
      Width           =   3555
   End
   Begin VB.ComboBox cmbEmp 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   960
      Sorted          =   -1  'True
      TabIndex        =   41
      Tag             =   "8"
      ToolTipText     =   "Select Disposition Code From List"
      Top             =   7320
      Width           =   1560
   End
   Begin VB.TextBox txtPCom 
      Height          =   285
      Left            =   5640
      TabIndex        =   40
      Tag             =   "1"
      Top             =   6600
      Width           =   915
   End
   Begin VB.TextBox txtLastOp 
      Height          =   285
      Left            =   840
      TabIndex        =   39
      Tag             =   "1"
      Top             =   6720
      Width           =   915
   End
   Begin VB.ListBox lstSelEmp 
      Height          =   1230
      Left            =   960
      Sorted          =   -1  'True
      TabIndex        =   38
      Top             =   8040
      Width           =   2175
   End
   Begin VB.CommandButton cmdAddEmp 
      Caption         =   "Add"
      Height          =   315
      Left            =   2640
      TabIndex        =   37
      Top             =   7320
      Width           =   915
   End
   Begin VB.CommandButton cmdDelEmp 
      Caption         =   "Delete"
      Height          =   315
      Left            =   3240
      TabIndex        =   36
      ToolTipText     =   "Cancel Selected Invoice"
      Top             =   8040
      Width           =   915
   End
   Begin VB.TextBox txtSCus 
      Height          =   285
      Left            =   6840
      TabIndex        =   34
      Tag             =   "1"
      Top             =   1440
      Width           =   915
   End
   Begin VB.CommandButton cmdNxt 
      Caption         =   "&Next >>"
      Enabled         =   0   'False
      Height          =   300
      Left            =   5196
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   120
      Width           =   732
   End
   Begin VB.CommandButton cmdLst 
      Caption         =   "<< &Last"
      Enabled         =   0   'False
      Height          =   300
      Left            =   4440
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   120
      Width           =   732
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "InspRTe01c.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   31
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox cmbRes 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1200
      Sorted          =   -1  'True
      TabIndex        =   1
      Tag             =   "8"
      ToolTipText     =   "Select Resposibility Code From List"
      Top             =   840
      Width           =   1675
   End
   Begin VB.ComboBox cmbDis 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1680
      Sorted          =   -1  'True
      TabIndex        =   9
      Tag             =   "8"
      ToolTipText     =   "Select Disposition Code From List"
      Top             =   4680
      Width           =   2040
   End
   Begin VB.ComboBox txtDte 
      Height          =   315
      Left            =   6720
      TabIndex        =   10
      Tag             =   "4"
      Top             =   4680
      Width           =   1095
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   7320
      Top             =   8880
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   9495
      FormDesignWidth =   7995
   End
   Begin VB.TextBox txtScr 
      Height          =   285
      Left            =   4440
      TabIndex        =   4
      Tag             =   "1"
      Top             =   1440
      Width           =   915
   End
   Begin VB.TextBox txtRwk 
      Height          =   285
      Left            =   3000
      TabIndex        =   3
      Tag             =   "1"
      Top             =   1440
      Width           =   915
   End
   Begin VB.TextBox txtQty 
      Height          =   285
      Left            =   1200
      TabIndex        =   2
      Tag             =   "1"
      Top             =   1440
      Width           =   915
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "&Delete"
      Height          =   315
      Left            =   6960
      TabIndex        =   23
      TabStop         =   0   'False
      ToolTipText     =   "Delete The Current Item"
      Top             =   960
      Width           =   875
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   315
      Left            =   6960
      TabIndex        =   19
      TabStop         =   0   'False
      ToolTipText     =   "Add An Item To The Report"
      Top             =   600
      Width           =   875
   End
   Begin VB.TextBox txtCor 
      Height          =   825
      Left            =   3960
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Tag             =   "9"
      Top             =   3720
      Width           =   3795
   End
   Begin VB.TextBox txtTst 
      Height          =   825
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Tag             =   "9"
      Top             =   3750
      Width           =   3555
   End
   Begin VB.TextBox txtDip 
      Height          =   1395
      Left            =   3960
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Tag             =   "9"
      Top             =   2010
      Width           =   3795
   End
   Begin VB.TextBox txtDis 
      Height          =   1395
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Tag             =   "9"
      Top             =   2010
      Width           =   3555
   End
   Begin VB.ComboBox cmbCde 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1200
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "8"
      ToolTipText     =   "Select Characteristic Code From List"
      Top             =   480
      Width           =   1675
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   6960
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Material Type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   12
      Left            =   3960
      TabIndex        =   49
      Top             =   5400
      Width           =   2445
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Revision"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   13
      Left            =   120
      TabIndex        =   48
      Top             =   5400
      Width           =   2445
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Employee"
      Height          =   285
      Index           =   15
      Left            =   120
      TabIndex        =   47
      Top             =   7320
      Width           =   825
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Percentage Complete"
      Height          =   285
      Index           =   16
      Left            =   3960
      TabIndex        =   46
      Top             =   6720
      Width           =   1605
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Last Op"
      Height          =   285
      Index           =   17
      Left            =   120
      TabIndex        =   45
      Top             =   6720
      Width           =   615
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Selected Work Centers"
      Height          =   285
      Index           =   18
      Left            =   120
      TabIndex        =   44
      Top             =   7800
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Submit to Customer"
      Height          =   165
      Index           =   14
      Left            =   5400
      TabIndex        =   35
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Responsibility"
      Height          =   405
      Index           =   11
      Left            =   120
      TabIndex        =   30
      Top             =   840
      Width           =   1185
   End
   Begin VB.Label lblRes 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3000
      TabIndex        =   29
      Top             =   840
      Width           =   3780
   End
   Begin VB.Label lblDis 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1680
      TabIndex        =   28
      Top             =   5040
      Width           =   2940
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Disposition Code"
      Height          =   405
      Index           =   10
      Left            =   120
      TabIndex        =   27
      Top             =   4680
      Width           =   1545
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Scrap"
      Height          =   285
      Index           =   9
      Left            =   3960
      TabIndex        =   26
      Top             =   1440
      Width           =   705
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Rework"
      Height          =   285
      Index           =   8
      Left            =   2280
      TabIndex        =   25
      Top             =   1440
      Width           =   765
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity"
      Height          =   285
      Index           =   7
      Left            =   120
      TabIndex        =   24
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label lblTag 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   285
      Left            =   1200
      TabIndex        =   22
      Top             =   120
      Width           =   1365
   End
   Begin VB.Label lblItem 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1"
      Height          =   285
      Left            =   3120
      TabIndex        =   21
      Top             =   120
      Width           =   375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Item"
      Height          =   285
      Index           =   6
      Left            =   2640
      TabIndex        =   20
      Top             =   120
      Width           =   615
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Corrective Action Date"
      Height          =   285
      Index           =   5
      Left            =   4800
      TabIndex        =   18
      Top             =   4680
      Width           =   1830
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Corrective Action"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   4
      Left            =   3960
      TabIndex        =   17
      Top             =   3480
      Width           =   2445
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Test and Investigation Results / Cause"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   3
      Left            =   120
      TabIndex        =   16
      Top             =   3480
      Width           =   3405
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Disposition Instructions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   2
      Left            =   3960
      TabIndex        =   15
      Top             =   1800
      Width           =   2445
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description of Discrepancy"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   14
      Top             =   1800
      Width           =   2445
   End
   Begin VB.Label lblCde 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3000
      TabIndex        =   13
      Top             =   480
      Width           =   3780
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Discrepancy"
      Height          =   405
      Index           =   0
      Left            =   120
      TabIndex        =   12
      Top             =   480
      Width           =   1185
   End
End
Attribute VB_Name = "InspRTe01c"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'4/18/06 Revisited. Changed FillCombo to eliminate blanks
Option Explicit
Dim RdoItems As ADODB.Recordset

Dim bOnLoad As Byte
Dim iItemIndex As Integer
Dim iItemCount As Integer
Dim iLastItem As Integer

Dim cQuantity As Currency
Dim sCompTag As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub cmbCde_Click()
   If Not bOnLoad Then GetCharaCode
   
End Sub


Private Sub cmbCde_LostFocus()
   cmbCde = CheckLen(cmbCde, 12)
   GetCharaCode
   On Error Resume Next
   sSql = "UPDATE RjitTable SET RITCHARCODE='" & Compress(cmbCde) & "' WHERE " _
          & "RITREF='" & sCompTag & "' AND RITITM=" & str(lblItem)
   clsADOCon.ExecuteSQL sSql
End Sub


Private Sub cmbDis_Click()
   GetDispCode
   
End Sub


Private Sub cmbDis_LostFocus()
   cmbDis = CheckLen(cmbDis, 12)
   GetDispCode
   On Error Resume Next
   sSql = "UPDATE RjitTable SET RITDISPCODE='" & Compress(cmbDis) & "' WHERE " _
          & "RITREF='" & sCompTag & "' AND RITITM=" & str(lblItem)
   clsADOCon.ExecuteSQL sSql
   
End Sub


Private Sub cmbRes_Click()
   GetRespCode
   
End Sub


Private Sub cmbRes_LostFocus()
   cmbRes = CheckLen(cmbRes, 12)
   GetRespCode
   On Error Resume Next
   sSql = "UPDATE RjitTable SET RITRESPCODE='" & Compress(cmbRes) & "' WHERE " _
          & "RITREF='" & sCompTag & "' AND RITITM=" & str(lblItem)
   clsADOCon.ExecuteSQL sSql
   
End Sub


Private Sub cmdAdd_Click()
   Additems
   
End Sub

Private Sub cmdCan_Click()
   Unload Me
   
End Sub


Private Sub cmdDel_Click()
   DeleteItem
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 6102
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub

Private Sub cmdLst_Click()
   If iItemIndex > 1 Then
      iItemIndex = iItemIndex - 1
      cmdNxt.Enabled = True
   Else
      iItemIndex = 1
      cmdLst.Enabled = False
      cmdNxt.Enabled = True
   End If
   ClearSelEmp
   GetCurrentItem
   
End Sub

Private Sub cmdNxt_Click()
   If iItemIndex < iItemCount Then
      iItemIndex = iItemIndex + 1
      cmdLst.Enabled = True
   Else
      iItemIndex = iItemCount
      cmdNxt.Enabled = False
      cmdLst.Enabled = True
   End If
   ClearSelEmp
   GetCurrentItem
   
End Sub

Private Sub cmdAddEmp_Click()
   Dim sItem As String
   
   Dim i As Integer
   Dim strEmp As String
   On Error Resume Next
   
   strEmp = Compress(cmbEmp)
   
   If (CheckIfEmpExists(strEmp) <> "") Then
      MsgBox "The Employee already exists in the List - " & strEmp & ".", _
         vbInformation, Caption
      Exit Sub
   End If
   
   ' Insert the part
   sSql = "INSERT INTO RjitEmpTable (RITREF,RITITM,PREMNUMBER) VALUES('" & sCompTag & "'," _
            & str(lblItem) & "," & strEmp & ")"
            
   clsADOCon.ExecuteSQL sSql ', rdExecDirect
   
   lstSelEmp.AddItem Format(strEmp, "000000")
   
   Exit Sub
DiaErr1:
   sProcName = "cmdAdd_Click"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub cmdDelEmp_Click()
   Dim sItem As String
   Dim i As Integer
   With lstSelEmp
      i = .ListIndex
      If i > -1 Then
         sItem = .List(i)
         On Error Resume Next
         
         sSql = "DELETE FROM RjitEmpTable WHERE " _
                  & "RITREF='" & sCompTag & "' AND RITITM=" & str(lblItem) _
                  & " AND PREMNUMBER = " & sItem
         
         clsADOCon.ExecuteSQL sSql ', rdExecDirect
         .RemoveItem (i)
         If i = .ListCount Then
            i = i - 1
         End If
         .ListIndex = i
      End If
   End With
   
   Exit Sub
DiaErr1:
   sProcName = "cmdDel_Click"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub


Private Function CheckIfEmpExists(strEmp As String) As String
   Dim RdoEmp As ADODB.Recordset
   On Error GoTo modErr1
   
   sSql = "SELECT PREMNUMBER FROM RjitEmpTable WHERE " _
            & " RITREF='" & sCompTag & "' AND RITITM=" & str(lblItem) _
            & " AND PREMNUMBER = " & strEmp
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoEmp, ES_FORWARD)
   If bSqlRows Then
      With RdoEmp
         CheckIfEmpExists = "" & Trim(!PREMNUMBER)
         ClearResultSet RdoEmp
      End With
   Else
      CheckIfEmpExists = ""

   End If
   Set RdoEmp = Nothing
   Exit Function

modErr1:
   sProcName = "CheckIfEmpExists"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MdiSect.ActiveForm

End Function

Private Sub ClearSelEmp()
   
   cmbEmp.Clear
   lstSelEmp.Clear
End Sub
Private Sub FillSelEmp()
   Dim RdoSelEmp As ADODB.Recordset
   On Error GoTo modErr1
   
   Dim sCompTag As String
   
   sCompTag = Compress(lblTag)
   
   sSql = "select PREMNUMBER from EmplTable where ((PREMTERMDT IS NULL) or (PREMTERMDT IS NOT NULL AND PREMREHIREDT > PREMTERMDT) ) AND ( PREMSTATUS NOT IN ('D','I'))" _
          & " order by PREMNUMBER"
   LoadNumComboBox cmbEmp, "000000"
   If bSqlRows Then cmbEmp = cmbEmp.List(0)
   
   sSql = "SELECT DISTINCT PREMNUMBER FROM RJitEmpTable WHERE RITREF = '" & Compress(lblTag) & "' " _
            & " AND RITITM = " & str(lblItem)
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoSelEmp, ES_FORWARD)
   If bSqlRows Then
      With RdoSelEmp
         Do Until .EOF
            lstSelEmp.AddItem "" & Format(Trim(.Fields(0)), "000000")
            .MoveNext
         Loop
         ClearResultSet RdoSelEmp
      End With
   End If
   Set RdoSelEmp = Nothing
   Exit Sub

modErr1:
   sProcName = "FillSelEmp"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MdiSect.ActiveForm

End Sub
Private Sub Form_Activate()
   MouseCursor 0
   If bOnLoad Then
      FillCombo
      If iItemCount = 0 Then GetItems
      bOnLoad = 0
   End If
   
End Sub

Private Sub Form_Load()
   SetFormSize Me
   FormatControls
   Move 200, 600
   iItemCount = 0
   bOnLoad = 1
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   InspRTe01b.optItm.Value = vbUnchecked
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   MdiSect.Enabled = True
   RdoItems.Close
   Set InspRTe01c = Nothing
   
End Sub



Private Sub GetItems()
   sCompTag = Compress(lblTag)
   MouseCursor 13
   On Error GoTo DiaErr1
   sSql = "SELECT RITREF,RITITM,RITCHARCODE,RITDATE FROM RjitTable WHERE RITREF='" _
          & sCompTag & "' ORDER BY RITITM"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoItems, ES_KEYSET)
   If bSqlRows Then
      With RdoItems
         lblItem = Format(!RITITM, "#0")
         cmbCde = "" & Trim(!RITCHARCODE)
         Do Until RdoItems.EOF
            iItemCount = iItemCount + 1
            iLastItem = !RITITM
            .MoveNext
         Loop
         ClearResultSet RdoItems
      End With
      GetCharaCode
      If iItemCount > 1 Then cmdNxt.Enabled = True
   Else
      RdoItems.AddNew
      RdoItems!RITREF = sCompTag
      RdoItems!RITITM = 1
      RdoItems.Update
'      bSqlRows = GetDataSet(RdoItems, ES_KEYSET)
'      If bSqlRows Then iLastItem = 1
      iLastItem = 1
      iItemCount = 1
   End If
   iItemIndex = 1
   GetCurrentItem
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "getitems"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub FillCombo()
   On Error GoTo DiaErr1
   sSql = "Qry_FillDescripancyCodes"
   LoadComboBox cmbCde
   
   sSql = "Qry_FillDispositionCodes"
   LoadComboBox cmbDis
   
   sSql = "Qry_FillReasonCodes"
   LoadComboBox cmbRes
   
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Private Sub Additems()
   Dim bResponse As Byte
   
   bResponse = MsgBox("Add Item " & str(iLastItem + 1) & "?", ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      MouseCursor 13
      On Error GoTo DiaErr1
      iLastItem = iLastItem + 1
      iItemCount = iItemCount + 1
      RdoItems.AddNew
      RdoItems!RITREF = sCompTag
      RdoItems!RITITM = iLastItem
      RdoItems!RITDATE = Null
      RdoItems.Update
      
      lblItem = str(iLastItem)
      cmbCde = ""
      txtDis = ""
      txtDip = ""
      txtCor = ""
      txtTst = ""
      txtDte = ""
      
      txtRev = ""
      txtMatType = ""
      txtSCus = ""
      txtLastOp = ""
      txtPCom = ""
      ClearSelEmp
      
      iItemIndex = iLastItem
      cmdLst.Enabled = True
      cmdNxt.Enabled = False
   Else
      CancelTrans
   End If
   On Error Resume Next
   MouseCursor 0
   cmbCde.SetFocus
   Exit Sub
   
DiaErr1:
   CurrError.Description = Err.Description
   MouseCursor 0
   MsgBox CurrError.Description & vbCr & " Can't Add Record.", vbExclamation, Caption
   
End Sub

Private Sub lblCde_Change()
   If Left(lblCde, 8) = "*** Char" Then
      lblCde.ForeColor = ES_RED
   Else
      lblCde.ForeColor = vbBlack
   End If
   
End Sub

Private Sub lblDis_Change()
   If Left(lblDis, 8) = "*** Disp" Then
      lblDis.ForeColor = ES_RED
   Else
      lblDis.ForeColor = vbBlack
   End If
   
End Sub

Private Sub lblRes_Change()
   If Left(lblDis, 8) = "*** Resp" Then
      lblRes.ForeColor = ES_RED
   Else
      lblRes.ForeColor = vbBlack
   End If
   
End Sub

Private Sub txtCor_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdLst_Click
   If KeyCode = vbKeyPageDown Then cmdNxt_Click
   
End Sub


Private Sub txtCor_LostFocus()
   On Error Resume Next
   txtCor = CheckLen(txtCor, 1020)
   txtCor = StrCase(txtCor, ES_FIRSTWORD)
   sSql = "UPDATE RjitTable SET RITCORA='" & txtCor & "' WHERE " _
          & "RITREF='" & sCompTag & "' AND RITITM=" & str(lblItem)
   clsADOCon.ExecuteSQL sSql
   
End Sub


Private Sub txtDip_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdLst_Click
   If KeyCode = vbKeyPageDown Then cmdNxt_Click
   
End Sub


Private Sub txtDip_LostFocus()
   On Error Resume Next
   txtDip = CheckLen(txtDip, 1020)
   txtDip = StrCase(txtDip, ES_FIRSTWORD)
   sSql = "UPDATE RjitTable SET RITDISP='" & txtDip & "' WHERE " _
          & "RITREF='" & sCompTag & "' AND RITITM=" & str(lblItem)
   clsADOCon.ExecuteSQL sSql
   
End Sub


Private Sub txtDis_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdLst_Click
   If KeyCode = vbKeyPageDown Then cmdNxt_Click
   
End Sub


Private Sub txtDis_LostFocus()
   On Error Resume Next
   txtDis = CheckLen(txtDis, 1020)
   txtDis = StrCase(txtDis, ES_FIRSTWORD)
   sSql = "UPDATE RjitTable SET RITDESC='" & txtDis & "' WHERE " _
          & "RITREF='" & sCompTag & "' AND RITITM=" & str(lblItem)
   clsADOCon.ExecuteSQL sSql
   
End Sub


Private Sub txtDte_DropDown()
   ShowCalendar Me
   
End Sub

Private Sub txtDte_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdLst_Click
   If KeyCode = vbKeyPageDown Then cmdNxt_Click
   
End Sub

Private Sub txtDte_LostFocus()
   On Error Resume Next
   If Len(Trim(txtDte)) > 0 Then
      If Trim(cmbDis) = "" Then
         MsgBox "Requires A Disposition Code.", _
            vbExclamation, Caption
         txtDte = ""
         Exit Sub
      End If
      txtDte = CheckDate(txtDte)
      sSql = "UPDATE RjitTable SET RITDATE='" & txtDte _
             & "',RITACT=1 WHERE " _
             & "RITREF='" & sCompTag & "' AND RITITM=" & str(lblItem)
   Else
      txtDte = ""
      sSql = "UPDATE RjitTable SET RITDATE=Null" _
             & ",RITACT=0 WHERE " _
             & "RITREF='" & sCompTag & "' AND RITITM=" & str(lblItem)
   End If
   clsADOCon.ExecuteSQL sSql
   
End Sub


Private Sub txtQty_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdLst_Click
   If KeyCode = vbKeyPageDown Then cmdNxt_Click
   
End Sub

Private Sub txtQty_LostFocus()
   txtQty = CheckLen(txtQty, 10)
   txtQty = Format(txtQty, ES_QuantityDataFormat)
   cQuantity = Val(txtQty)
   On Error Resume Next
   sSql = "UPDATE RjitTable SET RITQTY=" & txtQty & " WHERE " _
          & "RITREF='" & sCompTag & "' AND RITITM=" & str(lblItem)
   clsADOCon.ExecuteSQL sSql
   
End Sub


Private Sub txtRev_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdLst_Click
   If KeyCode = vbKeyPageDown Then cmdNxt_Click
   
End Sub

Private Sub txtRev_LostFocus()
   txtRev = CheckLen(txtRev, 1020)
   txtRev = StrCase(txtRev, ES_FIRSTWORD)
   sSql = "UPDATE RjitTable SET RITREV='" & txtRev & "' WHERE " _
          & "RITREF='" & sCompTag & "' AND RITITM=" & str(lblItem)
   clsADOCon.ExecuteSQL sSql
   
End Sub


Private Sub txtMatType_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdLst_Click
   If KeyCode = vbKeyPageDown Then cmdNxt_Click
   
End Sub

Private Sub txtMatType_LostFocus()
   txtMatType = CheckLen(txtMatType, 1020)
   txtMatType = StrCase(txtMatType, ES_FIRSTWORD)
   sSql = "UPDATE RjitTable SET RITMATTYPE='" & txtMatType & "' WHERE " _
          & "RITREF='" & sCompTag & "' AND RITITM=" & str(lblItem)
   clsADOCon.ExecuteSQL sSql
   
End Sub

Private Sub txtSCus_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdLst_Click
   If KeyCode = vbKeyPageDown Then cmdNxt_Click
   
End Sub

Private Sub txtSCus_LostFocus()
   txtSCus = CheckLen(txtSCus, 10)
   txtSCus = Format(txtSCus, ES_QuantityDataFormat)
   On Error Resume Next
   sSql = "UPDATE RjitTable SET RITCUSTQTY=" & txtSCus & " WHERE " _
          & "RITREF='" & sCompTag & "' AND RITITM=" & str(lblItem)
   
   clsADOCon.ExecuteSQL sSql
   
   
End Sub

Private Sub txtLastOp_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdLst_Click
   If KeyCode = vbKeyPageDown Then cmdNxt_Click
   
End Sub
Private Sub txtLastOp_LostFocus()
   On Error Resume Next
   
   If (IsNumeric(txtLastOp)) Then
      sSql = "UPDATE RjitTable SET RITLASTOP=" & txtLastOp & " WHERE " _
             & "RITREF='" & sCompTag & "' AND RITITM=" & str(lblItem)
      
      clsADOCon.ExecuteSQL sSql
   Else
      MsgBox ("Please enter a Operation Number")
   End If
   
End Sub

Private Sub txtPCom_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdLst_Click
   If KeyCode = vbKeyPageDown Then cmdNxt_Click
   
End Sub
Private Sub txtPCom_LostFocus()
   On Error Resume Next
   If (IsNumeric(txtPCom)) Then
      sSql = "UPDATE RjitTable SET RITPERCOMP=" & txtPCom & " WHERE " _
             & "RITREF='" & sCompTag & "' AND RITITM=" & str(lblItem)
      clsADOCon.ExecuteSQL sSql
   Else
      MsgBox ("Please enter a Percentage Complete.")
   End If
   
   
End Sub

Private Sub txtRwk_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdLst_Click
   If KeyCode = vbKeyPageDown Then cmdNxt_Click
   
End Sub

Private Sub txtRwk_LostFocus()
   txtRwk = CheckLen(txtRwk, 10)
   txtRwk = Format(txtRwk, ES_QuantityDataFormat)
   If Val(txtRwk) > cQuantity Then
      txtRwk = Format(cQuantity, ES_QuantityDataFormat)
      txtScr = Format(0, ES_QuantityDataFormat)
   End If
   On Error Resume Next
   sSql = "UPDATE RjitTable SET RITRWK=" & txtRwk & " WHERE " _
          & "RITREF='" & sCompTag & "' AND RITITM=" & str(lblItem)
   clsADOCon.ExecuteSQL sSql
   
End Sub

Private Sub txtScr_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdLst_Click
   If KeyCode = vbKeyPageDown Then cmdNxt_Click
   
End Sub

Private Sub txtScr_LostFocus()
   txtScr = CheckLen(txtScr, 10)
   txtScr = Format(txtScr, ES_QuantityDataFormat)
   If Val(txtScr) > cQuantity Then
      txtScr = Format(cQuantity, ES_QuantityDataFormat)
      txtRwk = Format(0, "#.000")
   End If
   If Val(txtScr) > (Val(txtRwk) + cQuantity) Then
      txtScr = Format(Val(txtQty) - Val(txtRwk), ES_QuantityDataFormat)
   End If
   If (Val(txtScr) + Val(txtRwk)) > cQuantity Then
      txtScr = Format(Val(txtQty) - Val(txtRwk), ES_QuantityDataFormat)
   End If
   On Error Resume Next
   sSql = "UPDATE RjitTable SET RITSCRP=" & txtScr & " WHERE " _
          & "RITREF='" & sCompTag & "' AND RITITM=" & str(lblItem)
   clsADOCon.ExecuteSQL sSql
   
   
End Sub

Private Sub txtTst_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdLst_Click
   If KeyCode = vbKeyPageDown Then cmdNxt_Click
   
End Sub


Private Sub txtTst_LostFocus()
   On Error Resume Next
   txtTst = CheckLen(txtTst, 1020)
   txtTst = StrCase(txtTst, ES_FIRSTWORD)
   sSql = "UPDATE RjitTable SET RITINVS='" & Trim(txtTst) & "' WHERE " _
          & "RITREF='" & sCompTag & "' AND RITITM=" & Trim(lblItem)
   clsADOCon.ExecuteSQL sSql
   
End Sub



Private Sub GetCurrentItem()
   Dim RdoBlob As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT RITREF,RITITM,RITCHARCODE,RITDATE," _
          & "RITQTY,RITRWK,RITSCRP,RITDESC," _
          & "RITDISPCODE,RITRESPCODE,RITREV, RITMATTYPE, RITCUSTQTY, RITLASTOP,RITPERCOMP " _
          & "FROM RjitTable WHERE RITREF='" _
          & sCompTag & "' AND RITITM=" & str(iItemIndex)
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoItems, ES_KEYSET)
   If bSqlRows Then
      With RdoItems
         lblItem = Trim(str(!RITITM))
         txtDis = "" & Trim(!RITDESC)
         If Not IsNull(!RITDATE) Then
            txtDte = Format(!RITDATE, "mm/dd/yy")
         Else
            txtDte = ""
         End If
         If !RITITM = iItemCount Then cmdNxt.Enabled = False
         If !RITITM = 1 Then cmdLst.Enabled = False
         txtQty = Format(!RITQTY, ES_QuantityDataFormat)
         txtRwk = Format(!RITRWK, ES_QuantityDataFormat)
         txtScr = Format(!RITSCRP, ES_QuantityDataFormat)
         cQuantity = Val(txtQty)
         cmbCde = "" & Trim(!RITCHARCODE)
         cmbRes = "" & Trim(!RITRESPCODE)
         cmbDis = "" & Trim(!RITDISPCODE)
         
         txtRev = "" & Trim(!RITREV)
         txtMatType = "" & Trim(!RITMATTYPE)
         txtSCus = Format(!RITCUSTQTY, ES_QuantityDataFormat)
         txtLastOp = "" & Trim(!RITLASTOP)
         txtPCom = "" & Trim(!RITPERCOMP)
         
         ClearResultSet RdoItems
         GetCharaCode
         GetDispCode
         GetRespCode
      End With
      sSql = "SELECT RITREF,RITITM,RITINVS " _
             & "FROM RjitTable WHERE RITREF='" _
             & sCompTag & "' AND RITITM=" & str(iItemIndex)
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoBlob, ES_KEYSET)
      If bSqlRows Then
         With RdoBlob
            txtTst = "" & Trim(!RITINVS)
            ClearResultSet RdoBlob
         End With
      End If
      sSql = "SELECT RITREF,RITITM,RITDISP " _
             & "FROM RjitTable WHERE RITREF='" _
             & sCompTag & "' AND RITITM=" & str(iItemIndex)
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoBlob, ES_KEYSET)
      If bSqlRows Then
         With RdoBlob
            txtDip = "" & Trim(!RITDISP)
            ClearResultSet RdoBlob
         End With
      End If
      sSql = "SELECT RITREF,RITITM,RITCORA " _
             & "FROM RjitTable WHERE RITREF='" _
             & sCompTag & "' AND RITITM=" & str(iItemIndex)
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoBlob, ES_KEYSET)
      If bSqlRows Then
         With RdoBlob
            txtCor = "" & Trim(!RITCORA)
            ClearResultSet RdoBlob
         End With
      End If
   End If
   
   ' Fill Employee detail
   FillSelEmp
      
   Set RdoBlob = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getcurrentit"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub DeleteItem()
   Dim bResponse As Byte
   
   bResponse = MsgBox("Really Delete Item " & lblItem & "?", ES_NOQUESTION, Caption)
   If bResponse = vbNo Then Exit Sub
   On Error GoTo DiaErr1
   MouseCursor 13
   sSql = "DELETE FROM RjitTable WHERE "
   sSql = sSql & "RITREF='" & sCompTag & "' AND RITITM=" & str(lblItem)
   clsADOCon.ExecuteSQL sSql
   MsgBox "Item Deleted.", vbInformation, Caption
   If iItemCount = 1 Then
      Unload Me
   Else
      If iItemIndex > 1 Then
         cmdLst_Click
      Else
         cmdNxt_Click
      End If
   End If
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   CurrError.Description = Err.Description
   MouseCursor 0
   MsgBox CurrError.Description & vbCr & " Can't Delete Record.", vbExclamation, Caption
   
End Sub

Private Sub GetCharaCode()
   Dim RdoCha As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT CDEREF,CDENUM,CDEDESC FROM RjcdTable " _
          & "WHERE CDEREF='" & Compress(cmbCde) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCha, ES_FORWARD)
   If bSqlRows Then
      With RdoCha
         cmbCde = "" & Trim(!CDENUM)
         lblCde = "" & Trim(!CDEDESC)
         ClearResultSet RdoCha
      End With
   Else
      If Len(Trim(cmbCde)) > 0 Then lblCde = "*** Characteristic Wasn't Found ***"
   End If
   Set RdoCha = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getcharco"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub GetRespCode()
   Dim RdoRsp As ADODB.Recordset
   If Trim(cmbRes) = "" Then
      lblRes = ""
      Exit Sub
   End If
   On Error GoTo DiaErr1
   sSql = "SELECT RESREF,RESNUM,RESDESC FROM RjrsTable " _
          & "WHERE RESREF='" & Compress(cmbRes) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRsp, ES_FORWARD)
   If bSqlRows Then
      With RdoRsp
         cmbRes = "" & Trim(!RESNUM)
         lblRes = "" & Trim(!RESDESC)
         ClearResultSet RdoRsp
      End With
   Else
      lblRes = "*** Responsibility Wasn't Found ***"
   End If
   Set RdoRsp = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getrespco"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub GetDispCode()
   Dim RdoDos As ADODB.Recordset
   If Trim(cmbDis) = "" Then
      lblDis = ""
      Exit Sub
   End If
   On Error GoTo DiaErr1
   sSql = "SELECT DISREF,DISNUM,DISDESC FROM RjdsTable " _
          & "WHERE DISREF='" & Compress(cmbDis) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoDos, ES_FORWARD)
   If bSqlRows Then
      With RdoDos
         cmbDis = "" & Trim(!DISNUM)
         lblDis = "" & Trim(!DISDESC)
         ClearResultSet RdoDos
      End With
   Else
      lblDis = "*** Disposition Wasn't Found ***"
   End If
   Set RdoDos = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getdispco"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub
