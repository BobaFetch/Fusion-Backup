VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form TempBox
   BorderStyle = 3 'Fixed Dialog
   Caption = "Change Invoice Customer"
   ClientHeight = 2985
   ClientLeft = 45
   ClientTop = 330
   ClientWidth = 5970
   ClipControls = 0 'False
   ControlBox = 0 'False
   LinkTopic = "Form1"
   LockControls = -1 'True
   MaxButton = 0 'False
   MDIChild = -1 'True
   MinButton = 0 'False
   ScaleHeight = 2985
   ScaleWidth = 5970
   ShowInTaskbar = 0 'False
   Begin VB.CommandButton cmdUpd
      Caption = "&Update"
      Height = 300
      Left = 3720
      TabIndex = 2
      Top = 1200
      Width = 855
   End
   Begin VB.ComboBox cmbCst
      Height = 315
      Left = 1920
      Sorted = -1 'True
      TabIndex = 1
      Tag = "3"
      ToolTipText = "Select Or Enter Customer"
      Top = 1200
      Width = 1555
   End
   Begin VB.ComboBox cmbPrc
      Height = 315
      Left = 1920
      Sorted = -1 'True
      TabIndex = 0
      Tag = "1"
      ToolTipText = "Select Or Enter Invoice"
      Top = 720
      Width = 1275
   End
   Begin VB.CommandButton cmdCan
      Cancel = -1 'True
      Caption = "Close"
      Height = 435
      Left = 3720
      TabIndex = 4
      TabStop = 0 'False
      Top = 0
      Width = 875
   End
   Begin ResizeLibCtl.ReSize ReSize1
      Left = 4440
      Top = 1800
      _Version = 196615
      _ExtentX = 741
      _ExtentY = 741
      _StockProps = 0
      Enabled = -1 'True
      FormMinWidth = 0
      FormMinHeight = 0
      FormDesignHeight = 2985
      FormDesignWidth = 5970
   End
   Begin VB.Label Label1
      BackStyle = 0 'Transparent
      Caption = "Editing.."
      BeginProperty Font
      Name = "MS Sans Serif"
      Size = 8.25
      Charset = 0
      Weight = 700
      Underline = 0 'False
      Italic = -1 'True
      Strikethrough = 0 'False
      EndProperty
      ForeColor = &H00800000&
      Height = 255
      Left = 240
      TabIndex = 6
      Top = 1680
      Width = 2295
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Customer Nickname"
      Height = 255
      Index = 1
      Left = 240
      TabIndex = 5
      Top = 1200
      Width = 1695
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Invoice Number"
      Height = 255
      Index = 0
      Left = 240
      TabIndex = 3
      Top = 720
      Width = 1695
   End
End
Attribute VB_Name = "TempBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RdoInv As rdoResultset
Dim bOnload As Byte
Dim bGoodInv As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbPrc_Click()
   bGoodInv = GetInvoice
   
End Sub


Private Sub cmbPrc_LostFocus()
   cmbPrc = Format(Abs(Val(cmbPrc)), "000000")
   bGoodInv = GetInvoice()
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub



Private Sub cmdUpd_Click()
   On Error Resume Next
   If bGoodInv = 1 Then
      MouseCursor 13
      sSql = "UPDATE CihdTable SET INVCUST='" & Compress(cmbCst) _
             & "' WHERE INVNO=" & Val(cmbPrc) & " "
      RdoCon.Execute sSql, rdExecDirect
      Label1 = "Updated."
      Label1.Refresh
      Sleep 1000
      Label1 = "Editing."
      MouseCursor 0
   Else
      MsgBox "No Invoice Selected.", vbExclamation, _
         Caption
   End If
   On Error Resume Next
   cmbPrc.SetFocus
   
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnload Then
      FillCombo
      bOnload = False
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   SetDiaPos Me
   FormatControls
   sCurrForm = Caption
   bOnload = True
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set RdoInv = Nothing
   Set TempBox = Nothing
   
End Sub



Private Sub FormatControls()
   Dim i As Integer
   FormatFormControls Me
   
   On Error Resume Next
   ReDim txtGotFocus(UBound(ESI_txtGotFocus))
   For i = 0 To UBound(ESI_txtGotFocus)
      Set txtGotFocus(i) = ESI_txtGotFocus(i)
   Next
   
   ReDim txtKeyPress(UBound(ESI_txtKeyPress))
   For i = 0 To UBound(ESI_txtKeyPress)
      Set txtKeyPress(i) = ESI_txtKeyPress(i)
   Next
   
   ReDim txtKeyDown(UBound(ESI_txtKeyDown))
   For i = 0 To UBound(ESI_txtKeyDown)
      Set txtKeyDown(i) = ESI_txtKeyDown(i)
   Next
   Erase ESI_txtGotFocus
   Erase ESI_txtKeyPress
   Erase ESI_txtKeyDown
   
End Sub

Public Sub FillCombo()
   Dim RdoCmb As rdoResultset
   On Error GoTo DiaErr1
   MouseCursor 13
   sSql = "SELECT INVNO FROM CihdTable"
   bSqlRows = GetDataSet(RdoCmb, ES_FORWARD)
   If bSqlRows Then
      With RdoCmb
         Do Until .EOF
            cmbPrc.AddItem Format(!INVNO, "000000")
            .MoveNext
         Loop
         .Cancel
      End With
   End If
   sSql = "SELECT CUREF,CUNICKNAME FROM CustTable"
   bSqlRows = GetDataSet(RdoCmb, ES_FORWARD)
   If bSqlRows Then
      With RdoCmb
         Do Until .EOF
            cmbCst.AddItem Trim(!CUNICKNAME)
            .MoveNext
         Loop
         .Cancel
      End With
   End If
   cmbPrc = cmbPrc.List(0)
   cmbCst = cmbCst.List(0)
   Set RdoCmb = Nothing
   bGoodInv = GetInvoice()
   Exit Sub
   
   DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Public Function GetInvoice() As Byte
   On Error Resume Next
   MouseCursor 13
   sSql = "SELECT INVNO,INVCUST FROM CihdTable " _
          & "WHERE INVNO=" & Val(cmbPrc) & " "
   bSqlRows = GetDataSet(RdoInv, ES_FORWARD)
   If bSqlRows Then
      With RdoInv
         cmbPrc = !INVNO
         cmbCst = "" & Trim(!INVCUST)
         .Cancel
      End With
      GetInvoice = 1
   Else
      GetInvoice = 0
   End If
   If Trim(cmbCst) <> "" Then GetCustomer
   MouseCursor 0
   
End Function

Public Sub GetCustomer()
   Dim RdoCst As rdoResultset
   On Error Resume Next
   sSql = "SELECT CUREF,CUNICKNAME FROM CustTable " _
          & "WHERE CUREF='" & Compress(cmbCst) & " "
   bSqlRows = GetDataSet(RdoCst, ES_FORWARD)
   If bSqlRows Then
      With RdoCst
         cmbCst = "" & Trim(!CUNICKNAME)
         .Cancel
      End With
   Else
      cmbCst = ""
   End If
   Set RdoCst = Nothing
   
End Sub
