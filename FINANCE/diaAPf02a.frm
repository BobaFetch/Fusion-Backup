VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaAPf02a
   BorderStyle = 3 'Fixed Dialog
   Caption = "Change AP Invoice GL Distribution"
   ClientHeight = 2280
   ClientLeft = 45
   ClientTop = 330
   ClientWidth = 6285
   ControlBox = 0 'False
   LinkTopic = "Form1"
   MaxButton = 0 'False
   MDIChild = -1 'True
   MinButton = 0 'False
   ScaleHeight = 2280
   ScaleWidth = 6285
   ShowInTaskbar = 0 'False
   Begin VB.CheckBox optFrm
      Height = 255
      Left = 1680
      TabIndex = 8
      Top = 120
      Visible = 0 'False
      Width = 735
   End
   Begin VB.CommandButton cmdDel
      Caption = "&Items"
      Height = 315
      Left = 5280
      TabIndex = 7
      ToolTipText = "Show Invoice Items"
      Top = 600
      Width = 875
   End
   Begin VB.ComboBox cmbInv
      Height = 315
      Left = 1560
      TabIndex = 0
      ToolTipText = "Vendor AP Invoices"
      Top = 960
      Width = 2055
   End
   Begin VB.CommandButton cmdCan
      Cancel = -1 'True
      Caption = "Close"
      Height = 435
      Left = 5280
      TabIndex = 1
      TabStop = 0 'False
      Top = 90
      Width = 875
   End
   Begin Threed.SSRibbon cmdHlp
      Height = 225
      Left = 0
      TabIndex = 2
      ToolTipText = "Subject Help"
      Top = 0
      Width = 255
      _Version = 65536
      _ExtentX = 450
      _ExtentY = 397
      _StockProps = 65
      BackColor = 12632256
      Autosize = 2
      RoundedCorners = 0 'False
      BevelWidth = 0
      Outline = 0 'False
      PictureUp = "diaAPf02a.frx":0000
      PictureDn = "diaAPf02a.frx":0146
   End
   Begin ResizeLibCtl.ReSize ReSize1
      Left = 5640
      Top = 1080
      _Version = 196615
      _ExtentX = 741
      _ExtentY = 741
      _StockProps = 0
      Enabled = -1 'True
      FormMinWidth = 0
      FormMinHeight = 0
      FormDesignHeight = 2280
      FormDesignWidth = 6285
   End
   Begin VB.Label lblNme
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      Height = 285
      Left = 1560
      TabIndex = 6
      Top = 1680
      Width = 3855
   End
   Begin VB.Label lblVnd
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      Height = 285
      Left = 1560
      TabIndex = 5
      Top = 1320
      Width = 1815
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Vendor"
      Height = 285
      Index = 0
      Left = 240
      TabIndex = 4
      Top = 1320
      Width = 1425
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Invoice Number"
      Height = 285
      Index = 3
      Left = 240
      TabIndex = 3
      Top = 960
      Width = 1425
   End
End
Attribute VB_Name = "diaAPf02a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001, ES/2002) is the property of                     ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***

'See the UpdateTables prodecure for database revisions

Option Explicit

'*************************************************************************************
'   diaAPf02a - Change AP Invoice GL Distributions
'
'   Notes:
'
'
'   Created: 11/13/02 (nth)
'   Revisons:
'
'
'*************************************************************************************

Dim bOnLoad As Byte
Dim bGoodInv As Boolean

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'*************************************************************************************

Private Sub cmbInv_Click()
   bGoodInv = GetInvoice()
   
End Sub

Private Sub cmbInv_LostFocus()
   cmbInv = CheckLen(cmbInv, 12)
   If cmbInv.ListCount > 0 Then
      If Len(Trim(cmbInv)) = 0 Then cmbInv = cmbInv.List(0)
   End If
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdDel_Click()
   bGoodInv = GetInvoice()
   If bGoodInv Then
      optFrm.Value = vbChecked
      'diaPsina.Show
   Else
      MsgBox "That Invoice Wasn't Found.", vbExclamation, Caption
   End If
   
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
   SetDiaPos Me
   FormatControls
   sCurrForm = Caption
   bOnLoad = True
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set diaAPf02a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Public Sub FillCombo()
   Dim RdoInv As rdoResultset
   On Error GoTo DiaErr1
   sSql = "SELECT DISTINCT VITNO FROM ViitTable "
   bSqlRows = GetDataSet(RdoInv)
   If bSqlRows Then
      With RdoInv
         cmbInv = Trim(!VITNO)
         Do Until .EOF
            'cmbInv.AddItem Trim(!VITNO)
            AddComboStr cmbInv.hWnd, "" & Trim(!VITNO)
            .MoveNext
         Loop
      End With
   End If
   On Error Resume Next
   Set RdoInv = Nothing
   If cmbInv.ListCount > 0 Then bGoodInv = GetInvoice()
   Exit Sub
   
   DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Public Function GetInvoice() As Boolean
   Dim RdoInv As rdoResultset
   Dim sInvoice As String
   Dim sVendor As String
   sInvoice = Trim(cmbInv)
   On Error GoTo DiaErr1
   sSql = "SELECT DISTINCT VITNO,VITVENDOR,VEREF," _
          & "VENICKNAME,VECNAME FROM ViitTable,VndrTable " _
          & "WHERE VITNO='" & sInvoice & "' AND VITVENDOR=VEREF"
   bSqlRows = GetDataSet(RdoInv)
   If bSqlRows Then
      With RdoInv
         lblVnd = Trim(!VENICKNAME)
         lblnme = Trim(!VECNAME)
      End With
      GetInvoice = True
   Else
      lblVnd = "** Invalid Invoice **"
      lblnme = ""
      GetInvoice = False
   End If
   Set RdoInv = Nothing
   Exit Function
   
   DiaErr1:
   sProcName = "getinvoice"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function



Private Sub optFrm_Click()
   'never visible - checks to see if items is loaded
   
End Sub
