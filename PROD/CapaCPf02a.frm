VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form CapaCPf02a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Delete Shops"
   ClientHeight    =   1860
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6240
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1860
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "CapaCPf02a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "D&elete"
      Enabled         =   0   'False
      Height          =   315
      Left            =   5280
      TabIndex        =   1
      ToolTipText     =   "Press To Delete This Shop"
      Top             =   600
      Width           =   915
   End
   Begin VB.ComboBox cmbShp 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1440
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Contains Available Shops (No Work Centers Attached)"
      Top             =   810
      Width           =   1815
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5280
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   915
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   360
      Top             =   1560
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   1860
      FormDesignWidth =   6240
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1440
      TabIndex        =   5
      Top             =   1200
      Width           =   2895
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Shop"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   1185
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   1170
      Width           =   1125
   End
End
Attribute VB_Name = "CapaCPf02a"
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
Dim bGoodShop As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbShp_Click()
   bGoodShop = GetShop()
   
End Sub

Private Sub cmbShp_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub cmbShp_LostFocus()
   Dim b As Byte
   Dim iList As Integer
   On Error Resume Next
   If cmbShp.ListCount > 0 Then
      For iList = 0 To cmbShp.ListCount - 1
         If cmbShp = cmbShp.List(iList) Then b = 1
      Next
      If b = 0 Then cmbShp = cmbShp.List(0)
      bGoodShop = GetShop()
   End If
   
End Sub

Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdDel_Click()
   If bGoodShop Then DeleteShop
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 4251
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub


Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      FillShops
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST
   FormatControls
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set CapaCPf02a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub FillShops()
   On Error GoTo DiaErr1
   cmbShp.Clear
   sSql = "SELECT DISTINCT SHPREF,SHPNUM FROM " _
          & "ShopTable LEFT JOIN WcntTable ON ShopTable.SHPREF=WcntTable.WCNSHOP " _
          & "WHERE (WcntTable.WCNSHOP Is Null)"
   LoadComboBox cmbShp
   If cmbShp.ListCount > 0 Then
      cmbShp = cmbShp.List(0)
   Else
      lblDsc = "*** No Shops Available To Delete ***"
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "fillshops"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Function GetShop() As Byte
   Dim sGetShop As String
   Dim RdoShp As ADODB.Recordset
   On Error GoTo DiaErr1
   sGetShop = Compress(cmbShp)
   If Len(sGetShop) > 0 Then
      sSql = "Qry_GetShop '" & sGetShop & "'"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoShp)
      If bSqlRows Then
         With RdoShp
            cmbShp = "" & Trim(!SHPNUM)
            lblDsc = "" & Trim(!SHPDESC)
            ClearResultSet RdoShp
         End With
         cmdDel.Enabled = True
         GetShop = True
      Else
         lblDsc = "*** Shop Doesn't Qualify ***"
         cmdDel = False
         GetShop = False
      End If
   End If
   Set RdoShp = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getshop"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub lblDsc_Change()
   If Left(lblDsc, 8) = "*** No S" Then
      lblDsc.ForeColor = ES_RED
      Exit Sub
   Else
      lblDsc.ForeColor = Es_TextForeColor
   End If
   If Left(lblDsc, 8) = "*** Shop" Then
      lblDsc.ForeColor = ES_RED
   Else
      lblDsc.ForeColor = Es_TextForeColor
   End If
   
End Sub


Private Sub DeleteShop()
   Dim bResponse As Byte
   Dim sShop As String
   Dim sMsg As String
   
   On Error GoTo DiaErr1
   sShop = Compress(cmbShp)
   sMsg = "Are You Certain That You Want To Delete" & vbCr _
          & "The Shop " & cmbShp & "?"
   bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
   If bResponse = vbYes Then
      sSql = "DELETE FROM ShopTable WHERE " _
             & "SHPREF='" & sShop & "' "
      clsADOCon.ExecuteSQL sSql
      If clsADOCon.RowsAffected Then
         MsgBox "The Shop Was Deleted.", _
            vbInformation, Caption
         FillShops
      Else
         MsgBox "Couldn't Delete The Shop.", _
            vbExclamation, Caption
      End If
   Else
      CancelTrans
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "deleteshop"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub
