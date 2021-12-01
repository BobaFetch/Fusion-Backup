VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form CapaCPf03a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Delete Work Centers"
   ClientHeight    =   2925
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
   ScaleHeight     =   2925
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "CapaCPf03a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox cmbWcn 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1440
      TabIndex        =   1
      Tag             =   "8"
      ToolTipText     =   "Contains Available Work Centers"
      Top             =   1560
      Width           =   1815
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "D&elete"
      Enabled         =   0   'False
      Height          =   315
      Left            =   5280
      TabIndex        =   2
      ToolTipText     =   "Press To Delete This Work Center"
      Top             =   600
      Width           =   915
   End
   Begin VB.ComboBox cmbShp 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1440
      TabIndex        =   0
      Tag             =   "8"
      ToolTipText     =   "Contains Available Shops (Work Centers Attached)"
      Top             =   810
      Width           =   1815
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5280
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   915
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   600
      Top             =   2640
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2925
      FormDesignWidth =   6240
   End
   Begin VB.Label lblStatus 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   2400
      Width           =   3495
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   285
      Index           =   3
      Left            =   240
      TabIndex        =   9
      Top             =   1920
      Width           =   1125
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Work Center"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   8
      Top             =   1590
      Width           =   1185
   End
   Begin VB.Label lblCenter 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1440
      TabIndex        =   7
      Top             =   1950
      Width           =   2895
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1440
      TabIndex        =   6
      Top             =   1200
      Width           =   2895
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Shop"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   840
      Width           =   1185
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   4
      Top             =   1170
      Width           =   1125
   End
End
Attribute VB_Name = "CapaCPf03a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'New 3/25/04
Option Explicit
Dim RdoCts As ADODB.Recordset
Dim bOnLoad As Byte
Dim bGoodCntr As Byte
Dim bGoodShop As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Function GetCenter() As Byte
   Dim RdoWcn As ADODB.Recordset
   Dim sCenter As String
   sSql = "Qry_GetRoutCenter '" & Compress(cmbWcn) & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoWcn, ES_FORWARD)
   If bSqlRows Then
      With RdoWcn
         cmbWcn = "" & Trim(!WCNNUM)
         lblCenter = "" & Trim(!WCNDESC)
         GetCenter = 1
         ClearResultSet RdoWcn
      End With
   Else
      GetCenter = 0
      lblCenter = "*** Work Center Wasn't Found."
   End If
   Set RdoWcn = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getcenter"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

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

Private Sub cmbWcn_Click()
   bGoodCntr = GetCenter()
   
End Sub


Private Sub cmbWcn_LostFocus()
   Dim b As Byte
   Dim iList As Integer
   On Error Resume Next
   If cmbWcn.ListCount > 0 Then
      For iList = 0 To cmbWcn.ListCount - 1
         If cmbWcn = cmbWcn.List(iList) Then b = 1
      Next
      If b = 0 Then cmbShp = cmbShp.List(0)
      bGoodCntr = GetCenter()
   End If
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdDel_Click()
   If bGoodCntr Then
      DeleteCenter
   Else
      MsgBox "Please Select A Valid Work Center.", vbInformation, Caption
   End If
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 4252
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
   Set RdoCts = Nothing
   Set CapaCPf03a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   lblStatus.ForeColor = ES_BLUE
   
End Sub

Private Sub FillShops()
   On Error GoTo DiaErr1
   cmbShp.Clear
   sSql = "SELECT DISTINCT SHPREF,SHPNUM FROM ShopTable " _
          & "LEFT JOIN WcntTable ON ShopTable.SHPREF=WcntTable.WCNSHOP " _
          & "WHERE (WcntTable.WCNSHOP Is Not Null)"
   LoadComboBox cmbShp
   If cmbShp.ListCount > 0 Then
      cmbShp = cmbShp.List(0)
   Else
      lblDsc = "*** No Shops With Work Centers Have Been Recorded."
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
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoShp, ES_FORWARD)
      If bSqlRows Then
         With RdoShp
            cmbShp = "" & Trim(!SHPNUM)
            lblDsc = "" & Trim(!SHPDESC)
            ClearResultSet RdoShp
         End With
         cmdDel.Enabled = True
         GetShop = 1
      Else
         lblDsc = "*** Shop Doesn't Qualify ***"
         cmdDel.Enabled = False
         GetShop = 0
      End If
   End If
   If GetShop = 1 Then FillCenters
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


Private Sub DeleteCenter()
   Dim bResponse As Byte
   Dim iBids As Integer
   Dim iLibrary As Integer
   Dim iMoOPs As Integer
   Dim iRoutings As Integer
   
   Dim sCenter As String
   Dim sMsg As String
   
   On Error GoTo DiaErr1
   clsADOCon.ADOErrNum = 0
   
   sCenter = Compress(cmbWcn)
   sMsg = "Delete work center " & cmbWcn & " in shop " & Me.cmbShp & "?"
   bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
   If bResponse = vbYes Then
      MouseCursor 13
      lblStatus.Visible = True
      sProcName = "checkLibrary"
      iLibrary = CheckLibrary()
      sProcName = "checkroutings"
      iRoutings = CheckRoutings()
      sProcName = "checkmoops"
      iMoOPs = CheckMoOPs()
      sProcName = "checkbids"
      iBids = CheckBids()
      MouseCursor 0
      sMsg = "This Work Center Is Included:       " & vbCr _
             & "On " & str$(iRoutings) & " Routing(s) " & vbCr _
             & "On " & str$(iLibrary) & " Routing Library Entrie(s)" & vbCr _
             & "On " & str$(iMoOPs) & " Manufacturing Order(s)" & vbCr _
             & "On " & str$(iBids) & " Estimate(s) " & vbCr _
             & "And Cannot Be Deleted."
      iBids = iBids + iMoOPs + iRoutings + iLibrary
      If iBids > 0 Then
         MsgBox sMsg, vbInformation, Caption
      Else
         'On Error Resume Next
         sSql = "DELETE FROM WcntTable WHERE WCNREF='" _
                & Compress(cmbWcn) & "' and WCNSHOP = '" & Me.cmbShp & "'"
         clsADOCon.ExecuteSQL sSql
         If clsADOCon.ADOErrNum = 0 Then
            MsgBox "Work center " & cmbWcn & " was successfully removed from shop " & cmbShp & ".", _
               vbInformation, Caption
            FillCenters
         Else
            MsgBox "Could Not Remove The Work Center.", _
               vbInformation, Caption
         End If
      End If
   Else
      CancelTrans
   End If
   lblStatus.Visible = False
   Exit Sub
   
DiaErr1:
   CurrError.Number = clsADOCon.ADOErrNum
   CurrError.Description = clsADOCon.ADOErrDesc
   DoModuleErrors Me
   
End Sub

Private Sub FillCenters()
   cmbWcn.Clear
   sSql = "Qry_FillWorkCenters '" & Compress(cmbShp) & "'"
   LoadComboBox cmbWcn
   If cmbWcn.ListCount > 0 Then
      cmbWcn = cmbWcn.List(0)
      bGoodCntr = GetCenter()
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "fillcenters"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Function CheckRoutings() As Integer
   lblStatus.Caption = "Checking Routings."
   sSql = "SELECT OPREF,OPCENTER FROM RtopTable WHERE " _
          & "OPSHOP='" & Compress(cmbShp) & "' AND " _
          & "OPCENTER='" & Compress(cmbWcn) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCts, ES_FORWARD)
   With RdoCts
      Do Until .EOF
         CheckRoutings = CheckRoutings + 1
         .MoveNext
      Loop
      ClearResultSet RdoCts
   End With
   
End Function

Private Function CheckMoOPs() As Integer
   lblStatus.Caption = "Checking Manufacturing Orders."
   sSql = "SELECT OPREF,OPCENTER FROM RnopTable WHERE " _
          & "OPSHOP='" & Compress(cmbShp) & "' AND " _
          & "OPCENTER='" & Compress(cmbWcn) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCts, ES_FORWARD)
   With RdoCts
      Do Until .EOF
         CheckMoOPs = CheckMoOPs + 1
         .MoveNext
      Loop
      ClearResultSet RdoCts
   End With
   
End Function

Private Function CheckBids() As Integer
   lblStatus.Caption = "Checking Estimates."
   sSql = "SELECT BIDRTEREF,BIDRTECENTER FROM EsrtTable WHERE " _
          & "BIDRTESHOP='" & Compress(cmbShp) & "' AND " _
          & "BIDRTECENTER='" & Compress(cmbWcn) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCts, ES_FORWARD)
   With RdoCts
      Do Until .EOF
         CheckBids = CheckBids + 1
         .MoveNext
      Loop
      ClearResultSet RdoCts
   End With
   
End Function

Private Function CheckLibrary() As Integer
   lblStatus.Caption = "Checking Estimates."
   sSql = "SELECT LIBREF,LIBCENTER FROM RlbrTable WHERE " _
          & "LIBSHOP='" & Compress(cmbShp) & "' AND " _
          & "LIBCENTER='" & Compress(cmbWcn) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCts, ES_FORWARD)
   With RdoCts
      Do Until .EOF
         CheckLibrary = CheckLibrary + 1
         .MoveNext
      Loop
      ClearResultSet RdoCts
   End With
   
End Function
