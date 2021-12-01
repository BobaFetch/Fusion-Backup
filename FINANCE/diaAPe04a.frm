VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaAPe04a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Revise Invoice Due Dates/Comments"
   ClientHeight    =   3300
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   5760
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H80000007&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3300
   ScaleWidth      =   5760
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbVnd 
      Height          =   315
      Left            =   1320
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Contains Vendors With Invoices"
      Top             =   360
      Width           =   1555
   End
   Begin VB.ComboBox cmbInv 
      Height          =   315
      Left            =   1320
      TabIndex        =   1
      ToolTipText     =   "List Of Invoices Found"
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   5535
   End
   Begin VB.ComboBox txtDue 
      Height          =   315
      Left            =   1320
      TabIndex        =   2
      Tag             =   "4"
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox txtCmt 
      Height          =   975
      Left            =   1320
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Tag             =   "9"
      Top             =   2160
      Width           =   4335
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   4320
      Top             =   0
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3300
      FormDesignWidth =   5760
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   4800
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   5
      ToolTipText     =   "Subject Help"
      Top             =   0
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   397
      _StockProps     =   65
      BackColor       =   12632256
      GroupAllowAllUp =   -1  'True
      Autosize        =   2
      RoundedCorners  =   0   'False
      BevelWidth      =   0
      Outline         =   0   'False
      PictureUp       =   "diaAPe04a.frx":0000
      PictureDn       =   "diaAPe04a.frx":0146
   End
   Begin VB.Label lblNme 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1320
      TabIndex        =   13
      Top             =   720
      Width           =   2775
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Vendor"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   12
      Top             =   360
      Width           =   945
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Invoices"
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   11
      Top             =   1080
      Width           =   945
   End
   Begin VB.Label lblCnt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5160
      TabIndex        =   10
      Top             =   1080
      Width           =   405
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Invoices Found"
      Height          =   285
      Index           =   10
      Left            =   3840
      TabIndex        =   9
      Top             =   1080
      Width           =   1305
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Due Date"
      Height          =   285
      Index           =   12
      Left            =   240
      TabIndex        =   7
      Top             =   1680
      Width           =   1065
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice Comments:"
      Height          =   405
      Index           =   11
      Left            =   240
      TabIndex        =   6
      Top             =   2160
      Width           =   945
   End
End
Attribute VB_Name = "diaAPe04a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of                     ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***

'See the UpdateTables prodecure for database revisions

Option Explicit

'*********************************************************************************
' diaAPe04a - Invoice Due Date / Comments
'
' Notes:
'
' Created: 12/27/02 (nth)
' Revisions:
'
'
'*********************************************************************************

'1 = true
'0 = false
Dim bOnLoad As Byte
Dim bCancel As Byte
Dim bGoodVendor As Byte
Dim bGoodInv As Byte
Dim AdoQry As ADODB.Command
Dim AdoParameter As ADODB.Parameter

Dim RdoInv As ADODB.Recordset

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'*********************************************************************************

Private Sub cmbInv_Click()
   If bCancel <> 1 Then
      cmbInv = CheckLen(cmbInv, 20)
      GetCmtAndDue
   End If
End Sub

Private Sub cmbInv_LostFocus()
   Dim bByte As Boolean
   Dim i As Integer
   
   If bCancel <> 1 Then
      cmbInv = CheckLen(cmbInv, 20)
      GetCmtAndDue
      
      'For i = 0 To cmbInv.ListCount - 1
      '    If cmbInv = cmbInv.List(i) Then bByte = True
      'Next
      'If Not bByte Then
      '    Beep
      '    If cmbInv.ListCount > 0 Then cmbInv = cmbInv.List(0)
      'End If
   End If
End Sub

Private Sub cmbVnd_Click()
   bGoodVendor = FindVendor(Me)
   If bGoodVendor Then
      GetInvoices
      GetCmtAndDue
   End If
End Sub

Private Sub cmbVnd_LostFocus()
   If bCancel <> 1 Then
      cmbVnd = CheckLen(cmbVnd, 10)
      If Len(cmbVnd) Then
         bGoodVendor = FindVendor(Me)
         If bGoodVendor Then
            GetInvoices
            GetCmtAndDue
         End If
      End If
   End If
End Sub

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, _
                             X As Single, Y As Single)
   bCancel = 1
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      SelectHelpTopic Me, "Revise Invoice Due Dates/Comments"
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub FillCombo()
   Dim RdoVed As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT DISTINCT VIVENDOR,VEREF,VENICKNAME " _
          & "FROM VihdTable,VndrTable WHERE VIVENDOR=VEREF"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoVed)
   If bSqlRows Then
      With RdoVed
         cmbVnd = "" & Trim(!VENICKNAME)
         Do Until .EOF
            
            AddComboStr cmbVnd.hWnd, "" & Trim(!VENICKNAME)
            .MoveNext
         Loop
         .Cancel
      End With
   End If
   Set RdoVed = Nothing
   If cmbVnd.ListCount > 0 Then
      bGoodVendor = FindVendor(Me)
      GetInvoices
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad = 1 Then
      FillCombo
      cmbVnd = cUR.CurrentVendor
      bGoodVendor = FindVendor(Me)
      bOnLoad = 0
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   Dim i As Integer
   FormLoad Me
   FormatControls
   sSql = "SELECT DISTINCT VINO,VIVENDOR FROM " _
          & "VihdTable WHERE VIVENDOR= ? "
   Set AdoQry = New ADODB.Command
   AdoQry.CommandText = sSql
   Set AdoParameter = New ADODB.Parameter
   AdoParameter.Type = adChar
   AdoParameter.SIZE = 10
   AdoQry.parameters.Append AdoParameter
   
   sCurrForm = Caption
   bOnLoad = 1
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If bGoodVendor Then
      cUR.CurrentVendor = cmbVnd
      SaveCurrentSelections
   End If
   Set RdoInv = Nothing
   Set AdoParameter = Nothing
   Set AdoQry = Nothing
   FormUnload
   Set diaAPe04a = Nothing
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Public Sub GetInvoices()
   Dim RdoInv As ADODB.Recordset
   Dim iTotal As Integer
   Dim sVendor As String
   
   On Error GoTo DiaErr1
   cmbInv.Clear
   sVendor = Compress(cmbVnd)
   AdoQry.parameters(0).Value = sVendor
   bSqlRows = clsADOCon.GetQuerySet(RdoInv, AdoQry)
   If bSqlRows Then
      With RdoInv
         Do Until .EOF
            iTotal = iTotal + 1
            AddComboStr cmbInv.hWnd, "" & Trim(!VINO)
            .MoveNext
         Loop
         .Cancel
      End With
   End If
   
   If cmbInv.ListCount > 0 Then
      cmbInv = cmbInv.List(0)
      GetCmtAndDue
   End If
   
   lblCnt = iTotal
   Set RdoInv = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getinvoices"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub GetCmtAndDue()
   On Error Resume Next
   RdoInv.Close
   On Error GoTo DiaErr1
   sSql = "SELECT VIDUEDATE,VICOMT FROM VihdTable WHERE VINO = '" _
          & cmbInv & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoInv, ES_KEYSET)
   If bSqlRows Then
      txtCmt.enabled = True
      txtDue.enabled = True
      txtCmt = "" & Trim(RdoInv!VICOMT)
      txtDue = Format("" & RdoInv!VIDUEDATE, "mm/dd/yy")
      bGoodInv = 1
   Else
      txtCmt.enabled = False
      txtDue.enabled = False
      txtCmt = ""
      txtDue = ""
      bGoodInv = 0
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "GetComAndDue"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub txtCmt_LostFocus()
   If bGoodInv = 1 Then
      txtCmt = CheckLen(txtCmt, 1020)
      txtCmt = CheckComments(txtCmt)
      On Error Resume Next
      With RdoInv
         !VICOMT = Trim(txtCmt)
         .Update
      End With
      If Err <> 0 Then ValidateEdit Me
   End If
End Sub

Private Sub txtDue_Click()
   txtDue = CheckDate(txtDue)
End Sub

Private Sub txtDue_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtDue_GotFocus()
   SelectFormat Me
End Sub

Private Sub txtDue_LostFocus()
   txtDue = CheckDate(txtDue)
   If bGoodInv = 1 Then
      On Error Resume Next
      With RdoInv
         !VIDUEDATE = txtDue
         .Update
      End With
      If Err <> 0 Then ValidateEdit Me
   End If
End Sub
