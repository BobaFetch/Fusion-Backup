VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaAPe05a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change Invoice GL Distribution"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5850
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   5850
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbVnd 
      Height          =   315
      Left            =   1515
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Contains Vendors With Invoices"
      Top             =   720
      Width           =   1555
   End
   Begin VB.ComboBox cmbInv 
      Height          =   315
      Left            =   1515
      TabIndex        =   1
      ToolTipText     =   "List Of Invoices Found"
      Top             =   1440
      Width           =   2055
   End
   Begin VB.CheckBox optFrm 
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "&Items"
      Height          =   315
      Left            =   4920
      TabIndex        =   2
      ToolTipText     =   "Show Invoice Items"
      Top             =   600
      Width           =   875
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   4920
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   90
      Width           =   875
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   4
      ToolTipText     =   "Subject Help"
      Top             =   0
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   397
      _StockProps     =   65
      BackColor       =   12632256
      Autosize        =   2
      RoundedCorners  =   0   'False
      BevelWidth      =   0
      Outline         =   0   'False
      PictureUp       =   "diaAPe05a.frx":0000
      PictureDn       =   "diaAPe05a.frx":0146
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   4080
      Top             =   0
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2520
      FormDesignWidth =   5850
   End
   Begin VB.Label lblNme 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1515
      TabIndex        =   10
      Top             =   1080
      Width           =   2775
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Vendor"
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   720
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Invoices"
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   8
      Top             =   1440
      Width           =   1305
   End
   Begin VB.Label lblCnt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5040
      TabIndex        =   7
      Top             =   1440
      Width           =   405
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Invoices Found"
      Height          =   285
      Index           =   10
      Left            =   3720
      TabIndex        =   6
      Top             =   1440
      Width           =   1185
   End
End
Attribute VB_Name = "diaAPe05a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of                     ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***

'See the UpdateTables prodecure for database revisions

'*************************************************************************************
' diaAPe05a - Change AP Invoice GL Distributions
'
' Notes:
'
' Created: 11/13/02 (nth)
' Revisons:
'   12/26/02 (nth) modified per JLH added vendor combo
'   09/30/03 (nth) fixed checklen from 12 to 20
'
'*************************************************************************************

Option Explicit

Dim bOnLoad As Byte
Dim bCancel As Byte
Dim bGoodInv As Byte
Dim bGoodVendor As Byte
Dim sMsg As String
Dim rdoQry As ADODB.Command
Dim AdoParameter1 As ADODB.Parameter


Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'*************************************************************************************

Private Sub cmbInv_Click()
   'If Not bCancel Then
   '  GetInvoices
   ' End If
End Sub

Private Sub cmbInv_LostFocus()
   If Not bCancel Then
      cmbInv = CheckLen(cmbInv, 20)
      If cmbInv.ListCount > 0 Then
         If Len(Trim(cmbInv)) = 0 Then cmbInv = cmbInv.List(0)
      End If
   End If
End Sub

Private Sub cmbVnd_Click()
   bGoodVendor = FindVendor(Me)
   If bGoodVendor Then GetInvoices
End Sub

Private Sub cmbVnd_LostFocus()
   cmbVnd = CheckLen(cmbVnd, 10)
   If Len(cmbVnd) Then
      bGoodVendor = FindVendor(Me)
      If bGoodVendor Then GetInvoices
   End If
End Sub

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, _
                             X As Single, Y As Single)
   bCancel = True
End Sub

Private Sub cmdDel_Click()
   If CheckJournal = 1 Then
      optFrm.Value = vbChecked
      diaAPe05b.Show
   End If
   
   'bGoodInv = GetInvoice()
   '   If bGoodInv Then
   
   'diaPsina.Show
   '       diaAPe05b.Show
   '  Else
   '     MsgBox "That Invoice Wasn't Found.", vbinformation, Caption
   'End If
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      SelectHelpTopic Me, "Change AP Invoice GL Distribution"
      cmdHlp = False
      MouseCursor 0
   End If
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillCombo
      bOnLoad = False
   End If
   If optFrm.Value = vbChecked Then
      'Unload diaPsina
      optFrm.Value = vbUnchecked
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   Dim i As Integer
   FormLoad Me
   FormatControls
   sCurrForm = Caption
'   sSql = "SELECT DISTINCT VINO,VIVENDOR FROM " _
'          & "VihdTable WHERE VIVENDOR= ? "

   sSql = "SELECT DISTINCT DCVENDORINV, DCVENDOR from JritTable where DCHEAD IN " _
            & " (SELECT DISTINCT MJGLJRNL from JrhdTable " _
            & " WHERE mjtype = 'PJ' and MJCLOSED IS NULL) AND DCVENDOR = ? "

   Set rdoQry = New ADODB.Command
   rdoQry.CommandText = sSql
   
   Set AdoParameter1 = New ADODB.Parameter
   AdoParameter1.Type = adChar
   AdoParameter1.SIZE = 10
   rdoQry.parameters.Append AdoParameter1
   
   bOnLoad = True
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If bGoodVendor Then
      cUR.CurrentVendor = cmbVnd
      SaveCurrentSelections
   End If
   FormUnload
   Set AdoParameter1 = Nothing
   Set rdoQry = Nothing
   Set diaAPe05a = Nothing
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Private Sub FillCombo()
   Dim RdoVed As ADODB.Recordset
   On Error GoTo DiaErr1
'   sSql = "SELECT DISTINCT VIVENDOR,VEREF,VENICKNAME " _
'          & "FROM VihdTable,VndrTable WHERE VIVENDOR=VEREF"

'   sSql = "SELECT DISTINCT DCVENDORINV,DCVENDOR from JritTable where DCHEAD IN " _
'            & " (SELECT DISTINCT MJGLJRNL from JrhdTable " _
'            & " WHERE mjtype = 'PJ' and MJCLOSED IS NULL) order by DCVENDOR"

   sSql = "SELECT DISTINCT VEREF,VENICKNAME " _
              & " From VndrTable, JritTable, JrhdTable " _
            & " Where DCVENDOR = VEREF " _
            & " AND DCHEAD = MJGLJRNL AND mjtype = 'PJ' and MJCLOSED IS NULL"


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


Public Sub GetInvoices()
   Dim RdoInv As ADODB.Recordset
   Dim iTotal As Integer
   Dim sVendor As String
   
   On Error GoTo DiaErr1
   cmbInv.Clear
   sVendor = Compress(cmbVnd)
   rdoQry.parameters(0).Value = sVendor
   
   bSqlRows = clsADOCon.GetQuerySet(RdoInv, rdoQry)
   If bSqlRows Then
      With RdoInv
         Do Until .EOF
            iTotal = iTotal + 1
            AddComboStr cmbInv.hWnd, "" & Trim(!DCVENDORINV)
            .MoveNext
         Loop
         .Cancel
      End With
   End If
   If cmbInv.ListCount > 0 Then cmbInv = cmbInv.List(0)
   lblCnt = iTotal
   Set RdoInv = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getinvoices"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub optFrm_Click()
   'never visible - checks to see if items is loaded
End Sub

Private Function CheckJournal() As Byte
   Dim rdoJrn As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT MJCLOSED, MJGLJRNL FROM JritTable INNER JOIN " _
          & "JrhdTable ON JritTable.DCHEAD = JrhdTable.MJGLJRNL " _
          & "WHERE (MJTYPE = 'PJ') AND (DCVENDORINV = '" & cmbInv & "') " _
          & "AND (DCVENDOR = '" & Compress(cmbVnd) & "')"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoJrn)
   If bSqlRows Then
      With rdoJrn
         If ("" & Trim(!MJCLOSED)) = "" Then
            CheckJournal = 1 'good
         Else
            sMsg = "Invoice " & cmbInv & " Resides In Closed Journal " _
                   & !MJGLJRNL & vbCrLf _
                   & "Reopen Before Revising GL Distributions."
            MsgBox sMsg, vbInformation, Caption
            CheckJournal = 0 'bad
         End If
      End With
   Else
      sMsg = "Invoice Journal Not Found."
      MsgBox sMsg, vbInformation, Caption
      CheckJournal = 0 'bad
   End If
   Set rdoJrn = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "checkjournal"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Function
