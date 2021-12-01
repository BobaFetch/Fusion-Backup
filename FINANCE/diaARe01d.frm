VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaARe01d 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Item To Invoice"
   ClientHeight    =   3375
   ClientLeft      =   1845
   ClientTop       =   1065
   ClientWidth     =   7080
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   7080
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCan 
      Caption         =   "&Cancel"
      Height          =   435
      Left            =   6120
      TabIndex        =   22
      Top             =   120
      Width           =   875
   End
   Begin VB.TextBox txtCmt 
      Height          =   975
      Left            =   1440
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Tag             =   "9"
      Top             =   2280
      Width           =   4335
   End
   Begin VB.TextBox txtDis 
      Height          =   285
      Left            =   1440
      TabIndex        =   5
      Tag             =   "1"
      Top             =   1920
      Width           =   615
   End
   Begin VB.CheckBox optCom 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   4080
      TabIndex        =   4
      Top             =   1920
      Width           =   735
   End
   Begin VB.TextBox txtFrt 
      Height          =   285
      Left            =   5880
      TabIndex        =   3
      Tag             =   "1"
      Top             =   1680
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtQty 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Tag             =   "1"
      ToolTipText     =   "Enter Quantity"
      Top             =   840
      Width           =   1095
   End
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   1440
      TabIndex        =   1
      Tag             =   "3"
      ToolTipText     =   "Select From List"
      Top             =   840
      Width           =   3075
   End
   Begin VB.TextBox txtPrc 
      Height          =   285
      Left            =   4680
      TabIndex        =   2
      Tag             =   "1"
      ToolTipText     =   "Price"
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton cmdAdd 
      Cancel          =   -1  'True
      Caption         =   "&Add"
      Height          =   315
      Left            =   6120
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   600
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5280
      Top             =   1200
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3375
      FormDesignWidth =   7080
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   21
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
      PictureUp       =   "diaARe01d.frx":0000
      PictureDn       =   "diaARe01d.frx":0146
   End
   Begin VB.Label z1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " %"
      Height          =   255
      Index           =   14
      Left            =   2040
      TabIndex        =   20
      Top             =   1920
      Width           =   375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Comments:"
      Height          =   255
      Index           =   13
      Left            =   240
      TabIndex        =   19
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Discount"
      Height          =   255
      Index           =   10
      Left            =   240
      TabIndex        =   18
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Pay Commission?:"
      Height          =   255
      Index           =   9
      Left            =   2640
      TabIndex        =   17
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label lblSon 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1680
      TabIndex        =   16
      Top             =   240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Freight Allow:"
      Height          =   255
      Index           =   8
      Left            =   5880
      TabIndex        =   15
      Top             =   2040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Item"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   14
      Top             =   240
      Width           =   615
   End
   Begin VB.Label lblItm 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   765
      TabIndex        =   13
      Top             =   240
      Width           =   375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity           "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   12
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number                                                "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   1440
      TabIndex        =   11
      Top             =   600
      Width           =   3135
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Price                 "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   4680
      TabIndex        =   10
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   1440
      TabIndex        =   9
      Top             =   1200
      Width           =   2895
   End
   Begin VB.Label lblRev 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1140
      TabIndex        =   8
      Top             =   240
      Width           =   255
   End
End
Attribute VB_Name = "diaARe01d"
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
' diaAre01d - Add item to invoice which was not originally on the sales order
'
' Created: 07/26/02 (nth)
' Revisions:
'
'
'*********************************************************************************

Dim AdoQry As ADODB.Command
Dim AdoParameter1 As ADODB.Parameter
Dim AdoParameter2 As ADODB.Parameter
Dim AdoParameter3 As ADODB.Parameter


Dim bOnLoad As Byte
Dim bCancel As Byte
Dim bGoodItem As Byte
Dim bGoodPart As Byte
Dim bNewItem As Byte

Dim iIndex As Integer
Dim iTotalItems As Integer
Dim iLastitem As Integer

Dim cOldQty As Currency

Dim lSalesOrder As Long
Dim vItem(300, 2) As Variant

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Private Sub cmbPrt_Click()
   bGoodPart = GetPart(cmbPrt)
End Sub

Private Sub cmbPrt_LostFocus()
   Dim sItm As String * 3
   cmbPrt = CheckLen(cmbPrt, 30)
   On Error Resume Next
   bGoodPart = GetPart(cmbPrt)
   If bGoodPart Then
      If Val(txtQty) > 0 Then
         sItm = Trim(lblItm)
         If Len(sItm) = 1 Then sItm = Chr$(32) & sItm
         sItm = sItm & lblRev
         
         'cmbJmp.List(iIndex - 1) = sItm & "-" & Left(cmbPrt, 10) & "... " & txtSdt
         'cmbJmp = cmbJmp.List(iIndex - 1)
         'cmbJmp.Enabled = True
         cmdAdd.enabled = True
         'cmdTrm.Enabled = True
         
         'z3.Enabled = True
         bNewItem = False
      End If
   End If
   
End Sub



Private Sub cmdAdd_Click()
   bCancel = 0
   UpdateItem
   Unload Me
End Sub

Private Sub cmdCan_Click()
   bCancel = 1
   Unload Me
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      SelectHelpTopic Me, "Sales Order Items"
      MouseCursor 0
      cmdHlp = False
   End If
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   
   If bOnLoad Then
      FillParts Me
      bOnLoad = False
      'lblFrd = diaCrvso.txtFrd
      GetAllItems
      'GetItems
      AddSoItem
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   SetFormSize Me
   Move diaARe01c.Left + 200, diaARe01c.Top + 200
   FormatControls
   
   lblSon = Format(0 + Val(Right(diaARe01c.lblSon, Len(diaARe01c.lblSon) - 1)), "#####0")
   
   
   lSalesOrder = Val(lblSon)
   sSql = "SELECT ITSO,ITNUMBER,ITREV,ITQTY," _
          & "ITPART,ITDOLLARS,ITSCHED,ITCUSTREQ,ITSCHEDDEL," _
          & "ITFRTALLOW,ITFRTALLOW,ITDISCRATE,ITCOMMISSION," _
          & "ITCOMMENTS FROM SoitTable WHERE ITSO= ? " _
          & "AND (ITNUMBER= ? AND ITREV= ? AND ITPSNUMBER='') "
   Set AdoQry = New ADODB.Command
   AdoQry.CommandText = sSql
   Set AdoParameter1 = New ADODB.Parameter
   AdoParameter1.Type = adInteger
   AdoQry.parameters.Append AdoParameter1
   
   Set AdoParameter2 = New ADODB.Parameter
   AdoParameter2.Type = adInteger
   AdoQry.parameters.Append AdoParameter2
   
   Set AdoParameter3 = New ADODB.Parameter
   AdoParameter3.Type = adChar
   AdoParameter3.SIZE = 2
   AdoQry.parameters.Append AdoParameter3
   
   
'   RdoQry.MaxRows = 1
   bOnLoad = True
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   On Error Resume Next
   If bCancel = 1 Then
      sSql = "DELETE FROM SoitTable WHERE ITSO=" & Val(lblSon) & " " _
             & "AND ITNUMBER=" & Val(lblItm) & " AND ITREV='" & lblRev & "' "
      clsADOCon.ExecuteSQL sSql
   End If
   diaARe01c.optItm = vbUnchecked
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   MdiSect.BotPanel = MdiSect.Caption
   On Error Resume Next
   'RdoRes.Close
   Set AdoParameter1 = Nothing
   Set AdoParameter2 = Nothing
   Set AdoParameter3 = Nothing
   
   Set AdoQry = Nothing
   Set diaARe01d = Nothing
End Sub


Public Function GetThisItem() As Byte
   Dim RdoItm As ADODB.Recordset
   Dim sItm As String * 2
   On Error GoTo DiaErr1
   'RdoQry.RowsetSize = 1
   'RdoQry(0) = Val(lblSon)
   'RdoQry(1) = Val(lblItm)
   'RdoQry(2) = lblRev
   AdoQry.parameters(0).Value = Val(lblSon)
   AdoQry.parameters(1).Value = Val(lblItm)
   AdoQry.parameters(2).Value = lblRev
   bSqlRows = clsADOCon.GetQuerySet(RdoItm, AdoQry, ES_DYNAMIC, True, 1)
   If bSqlRows Then
      With RdoItm
         lblItm = "" & str(!ITNUMBER)
         lblRev = "" & Trim(!itrev)
         'txtQty = Format(0 + !ITQTY, "####0.000")
         txtQty = !ITQTY
         cmbPrt = "" & Trim(!ITPART)
         'txtPrc = Format(0 + !ITDOLLARS, "####0.000")
         txtPrc = !ITDOLLARS
         txtFrt = Format(0 + !ITFRTALLOW, "####0.000")
         txtDis = Format(0 + !ITDISCRATE, "#0.000")
         If !ITCOMMISSION Then optCom.Value = vbChecked Else optCom.Value = vbUnchecked
         txtCmt = "" & RTrim(!ITCOMMENTS)
         'If Trim(txtRdt) = "" Then txtRdt = txtSdt
         'If Trim(txtDdt) = "" Then txtDdt = txtSdt
         .Cancel
      End With
      GetThisItem = True
      On Error Resume Next
      txtQty.SetFocus
   Else
      GetThisItem = False
   End If
   If GetThisItem Then
      bGoodPart = GetPart(cmbPrt)
      sItm = Trim(lblItm)
      If Len(sItm) = 1 Then sItm = Chr$(32) & sItm
      sItm = sItm & lblRev
      ' cmbJmp = sItm & "-" & Left(cmbPrt, 10) & "... " & txtSdt
   End If
   Set RdoItm = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getthisit"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Public Sub UpdateItem()
   Dim SPartRef
   If Val(txtQty) = 0 Then
      If Len(lblItm) > 0 Then
         MsgBox "Requires A Valid Quantity." & vbCrLf _
            & "Item Couldn't Be Updated.", vbInformation, Caption
      End If
      Exit Sub
   End If
   If Not bGoodPart Then
      MsgBox "Requires A Valid Part." & vbCrLf _
         & "Item Couldn't Be Updated.", vbInformation, Caption
      Exit Sub
   End If
   If Val(Trim(lblItm)) = 0 Then Exit Sub
   SPartRef = Compress(cmbPrt)
   
   'If Len(Trim(txtDdt)) = 0 Then txtDdt = txtSdt
   
   
   On Error GoTo DiaErr1
   sSql = "UPDATE SoitTable SET ITPART='" & SPartRef & "'," _
          & "ITQTY=" & Val(txtQty) & ",ITDOLLARS=" & Val(txtPrc) & "," _
          & "ITSCHED='" & Format(Now, "mm/dd/yyyy") & "',ITCUSTREQ='" & Format(Now, "mm/dd/yyyy") & "'," _
          & "ITSCHEDDEL='" & Format(Now, "mm/dd/yyyy") & "',ITFRTALLOW=" & Val(txtFrt) & "," _
          & "ITDISCRATE=" & Val(txtDis) & ",ITCOMMISSION=" & optCom.Value & "," _
          & "ITCOMMENTS='" & Trim(txtCmt) & "' WHERE ITSO=" & Val(lblSon) & " " _
          & "AND ITNUMBER=" & Val(lblItm) & " AND ITREV='" & lblRev & "' "
   clsADOCon.ExecuteSQL sSql
   If clsADOCon.RowsAffected = 0 Then
      MsgBox "Couldn't Update Item " & vbCrLf _
         & "Check Part, Dates, Comments.", vbInformation, Caption
   End If
   On Error Resume Next
   bNewItem = False
   'cmbJmp.Enabled = True
   cmdAdd.enabled = True
   'cmdTrm.Enabled = True
   Exit Sub
   
DiaErr1:
   sProcName = "updateitem"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Public Sub AddSoItem()
   Dim iNewItem As Integer
   Dim sNewDate As String
   Dim sNewPart As String
   
   'If Len(txtSdt) = 0 Then txtSdt = diaARe01b.txtDte
   'If Len(txtSdt) = 0 Then txtSdt = Format(Now, "mm/dd/yy")
   
   'If optRep.Value = vbUnchecked Then
   cmbPrt = ""
   txtPrc = "0.000"
   txtQty = "0.000"
   'End If
   
   sNewPart = Compress(cmbPrt)
   sNewDate = Format(Now, "mm/dd/yy")
   iNewItem = iLastitem + 1
   
   On Error GoTo DiaErr1
   sSql = "INSERT SoitTable (ITSO,ITNUMBER,ITPART,ITQTY,ITSCHED,ITBOOKDATE) " _
          & "VALUES(" & lblSon & "," & iNewItem & ",'" _
          & sNewPart & "'," & txtQty & ",'" & Format(Now, "mm/dd/yyyy") & "','" _
          & sNewDate & "')"
   clsADOCon.ExecuteSQL sSql
   
   If clsADOCon.RowsAffected > 0 Then
      iLastitem = iNewItem
      iTotalItems = iTotalItems + 1
      vItem(iTotalItems, 0) = iTotalItems
      vItem(iTotalItems, 1) = ""
      iIndex = iIndex + 1
      lblItm = "" & str(iLastitem)
      lblRev = ""
      txtCmt = ""
      
      'txtRdt = txtSdt
      'txtDdt = txtSdt
      'cmbJmp.AddItem Trim(lblItm) & "-" & Left(cmbPrt, 8) & "... " & txtSdt
      'cmbJmp = cmbJmp.List(cmbJmp.ListCount - 1)
      'cmbJmp.Enabled = False
      'cmdAdd.Enabled = False
      'cmdTrm.Enabled = False
      
      bNewItem = True
      
      'z3.Enabled = False
      
      On Error Resume Next
      txtQty.SetFocus
   Else
      MsgBox "Couldn't Add Item.", vbInformation, Caption
      On Error Resume Next
      bNewItem = False
      
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "addsoitem"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub lblDsc_Change()
   If Left(lblDsc, 8) = "*** Part" Then
      lblDsc.ForeColor = ES_RED
   Else
      lblDsc.ForeColor = vbBlack
   End If
   
End Sub

Private Sub lblSon_Click()
   'hold so number
End Sub

Public Function GetPart(sGetPart) As Byte
   Dim RdoPrt As ADODB.Recordset
   Dim sComment As String
   
   On Error GoTo DiaErr1
   sGetPart = Compress(sGetPart)
   If Len(sGetPart) > 0 Then
      sSql = "SELECT PARTREF,PARTNUM,PADESC,PAEXTDESC,PAPRICE FROM PartTable WHERE PARTREF='" & sGetPart & "'"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoPrt, ES_STATIC)
      If bSqlRows Then
         With RdoPrt
            cmbPrt = "" & Trim(!PARTNUM)
            lblDsc = "" & Trim(!PADESC) & vbCrLf
            sComment = "" & Trim(!PAEXTDESC)
            'If bNewItem Then txtPrc = Format(!PAPRICE, "####0.000")
            If bNewItem Then txtPrc = !PAPRICE
            'If Val(txtPrc) = 0 Then txtPrc = Format(!PAPRICE, "####0.000")
         End With
         GetPart = True
      Else
         GetPart = False
         cmbPrt = ""
         lblDsc = "*** Part Wasn't Found ***"
      End If
      On Error Resume Next
      Set RdoPrt = Nothing
   Else
      cmbPrt = ""
      lblDsc = ""
   End If
   lblDsc = lblDsc & sComment
   Set RdoPrt = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getpart"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Function

Private Sub optCom_KeyPress(KeyAscii As Integer)
'   KeyLock KeyAscii
End Sub

Private Sub optCom_LostFocus()
   If bGoodPart Then
      If Val(txtQty) > 0 Then
         'cmbJmp.Enabled = True
         cmdAdd.enabled = True
         'cmdTrm.Enabled = True
      End If
   End If
End Sub

Private Sub txtCmt_LostFocus()
   txtCmt = CheckLen(txtCmt, 255)
   txtCmt = StrCase(txtCmt, ES_FIRSTWORD)
   
   If bGoodPart Then
      If Val(txtQty) > 0 Then
         'cmbJmp.Enabled = True
         cmdAdd.enabled = True
         'cmdTrm.Enabled = True
      End If
   End If
End Sub






Private Sub txtDis_LostFocus()
   txtDis = CheckLen(txtDis, 6)
   txtDis = Format(Abs(Val(txtDis)), "#0.000")
   If bGoodPart Then
      If Val(txtQty) > 0 Then
         'cmbJmp.Enabled = True
         cmdAdd.enabled = True
         'cmdTrm.Enabled = True
      End If
   End If
   
End Sub

Private Sub txtFrt_LostFocus()
   txtFrt = CheckLen(txtFrt, 9)
   txtFrt = Format(Abs(Val(txtFrt)), "####0.000")
   If bGoodPart Then
      If Val(txtQty) > 0 Then
         'cmbJmp.Enabled = True
         cmdAdd.enabled = True
         'cmdTrm.Enabled = True
      End If
   End If
   
End Sub

Private Sub txtPrc_LostFocus()
   txtPrc = CheckLen(txtPrc, 9)
   txtPrc = Format(Abs(Val(txtPrc)), "####0.0000")
   If bGoodPart Then
      If Val(txtQty) > 0 Then
         'cmbJmp.Enabled = True
         cmdAdd.enabled = True
         'cmdTrm.Enabled = True
         'z3.Enabled = True
      End If
   End If
   
End Sub

Private Sub txtQty_GotFocus()
   cOldQty = Val(txtQty)
   
End Sub


Private Sub txtQty_LostFocus()
   txtQty = CheckLen(txtQty, 9)
   txtQty = Format(Abs(Val(txtQty)), "####0.0000")
   If Val(txtQty) = 0 Then txtQty = Format(cOldQty, "####0.0000")
   cOldQty = Val(txtQty)
   
End Sub


Public Sub GetAllItems()
   Dim RdoGet As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT ITSO,ITNUMBER FROM SoitTable " _
          & "WHERE ITSO=" & str(lSalesOrder) & " ORDER BY ITNUMBER DESC"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoGet)
   If bSqlRows Then iLastitem = Format(RdoGet!ITNUMBER, "##0")
   Set RdoGet = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getallitem"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub
