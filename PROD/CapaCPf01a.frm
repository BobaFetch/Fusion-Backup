VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form CapaCPf01a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reassign Shops And Work Centers"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6855
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "CapaCPf01a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   20
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CheckBox optDel 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2160
      TabIndex        =   7
      ToolTipText     =   "Delete This Center When Finished"
      Top             =   2640
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CheckBox optEst 
      Caption         =   "Estimates"
      Height          =   255
      Left            =   5640
      TabIndex        =   6
      ToolTipText     =   "Update Bids"
      Top             =   2220
      Width           =   1215
   End
   Begin VB.CommandButton cmdAsn 
      Caption         =   "Apply"
      Enabled         =   0   'False
      Height          =   315
      Left            =   5760
      TabIndex        =   8
      ToolTipText     =   "Update And Apply Changes"
      Top             =   2640
      Width           =   915
   End
   Begin VB.CheckBox optMo 
      Caption         =   "Manufacturing Orders"
      Height          =   255
      Left            =   3480
      TabIndex        =   5
      ToolTipText     =   "Update MO's"
      Top             =   2220
      Value           =   1  'Checked
      Width           =   2055
   End
   Begin VB.CheckBox optRte 
      Caption         =   "Routings"
      Height          =   255
      Left            =   2160
      TabIndex        =   4
      ToolTipText     =   "Update Routings"
      Top             =   2220
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.ComboBox cmbWcn 
      Height          =   288
      Index           =   1
      Left            =   2160
      Sorted          =   -1  'True
      TabIndex        =   3
      Tag             =   "3"
      ToolTipText     =   "Select Or Enter A New Work Center"
      Top             =   1800
      Width           =   1815
   End
   Begin VB.ComboBox cmbShp 
      ForeColor       =   &H00800000&
      Height          =   288
      Index           =   1
      Left            =   2160
      Sorted          =   -1  'True
      TabIndex        =   2
      Tag             =   "8"
      ToolTipText     =   "Select From List"
      Top             =   1440
      Width           =   1815
   End
   Begin VB.ComboBox cmbWcn 
      Height          =   288
      Index           =   0
      Left            =   2160
      Sorted          =   -1  'True
      TabIndex        =   1
      Tag             =   "3"
      ToolTipText     =   "Select Or Enter A New Work Center"
      Top             =   960
      Width           =   1815
   End
   Begin VB.ComboBox cmbShp 
      ForeColor       =   &H00800000&
      Height          =   288
      Index           =   0
      Left            =   2160
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "8"
      ToolTipText     =   "Select From List"
      Top             =   600
      Width           =   1815
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5880
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   0
      Width           =   915
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5760
      Top             =   3360
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3120
      FormDesignWidth =   6855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Delete The Work Center"
      Height          =   192
      Index           =   5
      Left            =   240
      TabIndex        =   19
      Top             =   2640
      Visible         =   0   'False
      Width           =   2052
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Update All of....."
      Height          =   192
      Index           =   4
      Left            =   240
      TabIndex        =   18
      Top             =   2220
      Width           =   1572
   End
   Begin VB.Label lblCenter 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Index           =   1
      Left            =   4080
      TabIndex        =   17
      Top             =   1800
      Width           =   2652
   End
   Begin VB.Label lblShop 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Index           =   1
      Left            =   4080
      TabIndex        =   16
      Top             =   1440
      Width           =   2652
   End
   Begin VB.Label lblCenter 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Index           =   0
      Left            =   4080
      TabIndex        =   15
      Top             =   960
      Width           =   2652
   End
   Begin VB.Label lblShop 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Index           =   0
      Left            =   4080
      TabIndex        =   14
      Top             =   600
      Width           =   2652
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "New Work Center"
      Height          =   252
      Index           =   3
      Left            =   240
      TabIndex        =   13
      Top             =   1800
      Width           =   1572
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "New Shop"
      Height          =   192
      Index           =   1
      Left            =   240
      TabIndex        =   12
      Top             =   1440
      Width           =   1572
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Existing Work Center"
      Height          =   252
      Index           =   0
      Left            =   240
      TabIndex        =   11
      Top             =   960
      Width           =   1572
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Existing Shop"
      Height          =   192
      Index           =   2
      Left            =   240
      TabIndex        =   10
      Top             =   600
      Width           =   1572
   End
End
Attribute VB_Name = "CapaCPf01a"
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

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd


Private Sub cmbShp_Click(Index As Integer)
   GetShop (Index)
   FillCenters Index
   
End Sub


Private Sub cmbShp_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub cmbShp_LostFocus(Index As Integer)
   Dim b As Byte
   Dim iList As Integer
   For iList = 0 To cmbShp(Index).ListCount - 1
      If cmbShp(Index) = cmbShp(Index).List(iList) Then b = 1
   Next
   If b = 0 Then
      Beep
      cmbShp(Index) = cmbShp(Index).List(0)
   End If
   GetShop Index
   FillCenters Index
   
End Sub


Private Sub cmbWcn_Click(Index As Integer)
   GetCenter Index
   
End Sub


Private Sub cmbWcn_GotFocus(Index As Integer)
   SelectFormat Me
   
End Sub

Private Sub cmbWcn_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyCase KeyAscii
   
End Sub


Private Sub cmbWcn_LostFocus(Index As Integer)
   Dim b As Byte
   Dim iList As Integer
   For iList = 0 To cmbWcn(Index).ListCount - 1
      If cmbWcn(Index) = cmbWcn(Index).List(iList) Then b = 1
   Next
   If b = 0 Then
      Beep
      cmbWcn(Index) = cmbWcn(Index).List(0)
   End If
   GetCenter Index
   
End Sub


Private Sub cmdAsn_Click()
   If optRte.Value = vbUnchecked And optMo.Value = vbUnchecked Then
      MsgBox "Select Either Routings, MO's Or Both.", _
         vbInformation, Caption
      Exit Sub
   End If
   If (Trim(cmbShp(0))) & (Trim(cmbWcn(0))) = (Trim(cmbShp(1))) & (Trim(cmbWcn(1))) Then
      MsgBox "The Work Centers Are The Same.", _
         vbExclamation, Caption
   Else
      UpdateCenters
   End If
   
End Sub

Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 4250
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub



Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      FillCombos
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
   Set CapaCPf01a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub optDel_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optEst_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optMo_Click()
   If optRte And optMo Then optDel.Enabled = True _
                                             Else optDel.Enabled = False
   
End Sub

Private Sub optMo_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optRte_Click()
   If optRte And optMo Then optDel.Enabled = True _
                                             Else optDel.Enabled = False
   
End Sub

Private Sub optRte_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub



Private Sub FillCombos()
   Dim RdoCmb As ADODB.Recordset
   On Error GoTo DiaErr1
   cmbShp(0).Clear
   cmbShp(1).Clear
   sSql = "Qry_FillShops "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCmb)
   If bSqlRows Then
      With RdoCmb
         Do Until .EOF
            AddComboStr cmbShp(0).hwnd, "" & Trim(!SHPNUM)
            AddComboStr cmbShp(1).hwnd, "" & Trim(!SHPNUM)
            .MoveNext
         Loop
         ClearResultSet RdoCmb
      End With
   End If
   Set RdoCmb = Nothing
   If cmbShp(0).ListCount > 0 Then
      cmbShp(0) = cmbShp(0).List(0)
      cmbShp(1) = cmbShp(0).List(0)
      GetShop 0
      GetShop 1
      FillCenters 0
      FillCenters 1
   End If
   On Error Resume Next
   cmbShp(0).SetFocus
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombos"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub FillCenters(iIndex As Integer)
   Dim RdoCmb As ADODB.Recordset
   cmbWcn(iIndex).Clear
   On Error GoTo DiaErr1
   '1/6/04
   sSql = "SELECT WCNREF,WCNNUM,WCNSHOP FROM WcntTable WHERE WCNSHOP='" & Compress(cmbShp(iIndex)) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCmb, ES_FORWARD)
   If bSqlRows Then
      With RdoCmb
         Do Until .EOF
            AddComboStr cmbWcn(iIndex).hwnd, "" & Trim(!WCNNUM)
            .MoveNext
         Loop
         ClearResultSet RdoCmb
      End With
   End If
   If cmbWcn(iIndex).ListCount > 0 Then
      cmbWcn(iIndex) = cmbWcn(iIndex).List(0)
      GetCenter 0
      GetCenter 1
   End If
   Set RdoCmb = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "fillcenters"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub GetShop(Index As Integer)
   Dim RdoShp As ADODB.Recordset
   Dim sShop As String
   sShop = Compress(cmbShp(Index))
   sSql = "Qry_GetShop '" & sShop & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoShp, ES_FORWARD)
   If bSqlRows Then
      lblShop(Index) = "" & Trim(RdoShp!SHPDESC)
   Else
      lblShop(Index) = ""
   End If
   Set RdoShp = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getshop"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub GetCenter(iIndex As Integer)
   Dim RdoWcn As ADODB.Recordset
   Dim b As Byte
   Dim sCenter As String
   sCenter = Compress(cmbWcn(iIndex))
   sSql = "Qry_GetRoutCenter '" & sCenter & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoWcn, ES_FORWARD)
   If bSqlRows Then
      cmbWcn(iIndex) = "" & Trim(RdoWcn!WCNNUM)
      lblCenter(iIndex) = "" & Trim(RdoWcn!WCNDESC)
      cmdAsn.Enabled = True
      b = 1
   Else
      lblCenter(iIndex) = ""
      cmdAsn.Enabled = False
      b = 0
   End If
   '        If b = 1 Then
   '            If cmbWcn(1) <> cmbWcn(0) Then cmdAsn.Enabled = True _
   '                Else cmdAsn.Enabled = False
   '        End If
   Set RdoWcn = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getcenter"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

'3/25/04 Removed the Center Delete and added Estimates

Private Sub UpdateCenters()
   Dim bResponse As Byte
   Dim lBids As Long
   Dim lLibrarys As Long
   Dim lMoOps As Long
   Dim lRoutings As Long
   Dim sOldShop As String
   Dim sOldCenter As String
   Dim sNewShop As String
   Dim sNewCenter As String
   Dim sMsg As String
   
   sOldShop = Compress(cmbShp(0))
   sOldCenter = Compress(cmbWcn(0))
   sNewShop = Compress(cmbShp(1))
   sNewCenter = Compress(cmbWcn(1))
   
   On Error GoTo DiaErr1
   sMsg = "You Have Chosen To Update The Following:  "
   If optRte.Value = vbChecked Then _
                     sMsg = sMsg & vbCr & "Routing Operations And Library Ops"
   If optMo.Value = vbChecked Then _
                    sMsg = sMsg & vbCr & "Manufacturing Order Operations"
   sMsg = sMsg & vbCr & "Do You Want To Continue?.."
   
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      cmdAsn.Enabled = False
      'On Error Resume Next
      MouseCursor 13
      If optMo.Value = vbChecked Then
         sSql = "UPDATE RnopTable SET OPSHOP='" & sNewShop & "'," _
                & "OPCENTER='" & sNewCenter & "' WHERE " _
                & "OPCENTER='" & sOldCenter & "' AND OPSHOP='" & sOldShop & "'"
         clsADOCon.ExecuteSQL sSql
         lMoOps = clsADOCon.RowsAffected
      End If
      If optEst.Value = vbChecked Then
         sSql = "UPDATE EsrtTable SET BIDRTESHOP='" & sNewShop & "'," _
                & "BIDRTECENTER='" & sNewCenter & "' WHERE " _
                & "BIDRTECENTER='" & sOldCenter & "' AND BIDRTESHOP='" & sOldShop & "'"
         clsADOCon.ExecuteSQL sSql
         lBids = clsADOCon.RowsAffected
      End If
      If optRte.Value = vbChecked Then
         sSql = "UPDATE RtopTable SET OPSHOP='" & sNewShop & "'," _
                & "OPCENTER='" & sNewCenter & "' WHERE " _
                & "OPCENTER='" & sOldCenter & "' AND OPSHOP='" & sOldShop & "'"
         clsADOCon.ExecuteSQL sSql
         lRoutings = clsADOCon.RowsAffected
         
         sSql = "UPDATE RlbrTable SET LIBSHOP='" & sNewShop & "'," _
                & "LIBCENTER='" & sNewCenter & "' WHERE " _
                & "LIBCENTER='" & sOldCenter & "' AND LIBSHOP='" & sOldShop & "'"
         clsADOCon.ExecuteSQL sSql
         lLibrarys = clsADOCon.RowsAffected
      End If
      MouseCursor 0
      sMsg = "MO Operatons Updated... " & Trim(str(lMoOps)) & vbCr _
             & "Routing Ops Updated.... " & Trim(str(lRoutings)) & vbCr _
             & "Library Ops Updated.... " & Trim(str(lLibrarys)) & vbCr _
             & "Estimate Ops Updated.... " & Trim(str(lBids))
      MsgBox sMsg, vbInformation, Caption
      
      '            If optDel.Value = vbChecked Then
      '                sMsg = "You Have Chosen To Remove This Work Center." & vbCr _
      '                    & "Do You Want To Continue And Delete It?"
      '                bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
      '                    If bResponse = vbYes Then
      '                        sSql = "DELETE FROM WcntTable WHERE " _
      '                            & "WCNREF='" & sOldCenter & "'"
      '                        clsAdoCon.ExecuteSQL sSql
      '                        If clsadocon.rowsaffected > 0 Then
      '                            MsgBox "Work Center Deleted.", _
      '                                vbInformation, Caption
      '                            FillCombos
      '                        Else
      '                            MsgBox "Couldn't Delete Work Center.", _
      '                                vbExclamation, Caption
      '                        End If
      '                     End If
      '            End If
   Else
      CancelTrans
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "updatecenters"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub
