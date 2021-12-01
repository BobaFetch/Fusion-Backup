VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaAPf01a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cancel An AP Invoice"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5880
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   5880
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbVnd 
      Height          =   315
      Left            =   960
      Sorted          =   -1  'True
      TabIndex        =   7
      Tag             =   "3"
      ToolTipText     =   "Contains Vendors With PO's Not Invoiced"
      Top             =   360
      Width           =   1555
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "C&ancel"
      Height          =   315
      Left            =   4920
      TabIndex        =   6
      ToolTipText     =   "Cancel This Invoice"
      Top             =   600
      Width           =   875
   End
   Begin VB.ComboBox cmbInv 
      Height          =   315
      Left            =   960
      TabIndex        =   0
      ToolTipText     =   "Includes Invoices Without Cash Disbursements"
      Top             =   1200
      Width           =   2775
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   4920
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   90
      Width           =   875
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   2
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
      PictureUp       =   "diaAPf01a.frx":0000
      PictureDn       =   "diaAPf01a.frx":0146
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5400
      Top             =   960
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2175
      FormDesignWidth =   5880
   End
   Begin VB.Label lblFnd 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   4560
      TabIndex        =   9
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Found"
      Height          =   285
      Index           =   1
      Left            =   3960
      TabIndex        =   8
      Top             =   1200
      Width           =   585
   End
   Begin VB.Label lblNme 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   960
      TabIndex        =   5
      Top             =   720
      Width           =   2775
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Vendor"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   360
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice"
      Height          =   285
      Index           =   3
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   705
   End
End
Attribute VB_Name = "diaAPf01a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of                     ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***

' See the UpdateTables prodecure for database revisions

'*****************************************************************************************
'
' diaAPf01a - Cancel AP Invoice
'
' Created: (nth)
' Revisions:
' 12/26/02 (nth) Look for payment in check setup before canceling per JLH.
' 02/04/03 (nth) Check for closed MO and close journal per JDA.
' 05/29/03 (nth) Released invoice type 17 back to type 15 per SJW
' 01/27/05 (nth) Remove commissions reference.
'
'*****************************************************************************************

Option Explicit

Dim bOnLoad As Byte
Dim bCancel As Byte
Dim bGoodInv As Byte

Dim sJournalID As String
Dim sMsg As String
Dim sVendor As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'*****************************************************************************************

Private Sub cmbVnd_Click()
   FindVendor Me
   GetInvoices
End Sub

Private Sub cmbVnd_LostFocus()
   FindVendor Me
   GetInvoices
End Sub

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, _
                             x As Single, y As Single)
   bCancel = True
End Sub

Private Sub cmdDel_Click()
   If CheckInvoice Then
      CancelInvoice
   End If
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      SelectHelpTopic Me, "Cancel An AP Invoice"
      cmdHlp = False
      MouseCursor 0
   End If
End Sub

Private Sub Form_Activate()
   
   MdiSect.lblBotPanel = Caption
   
   If bOnLoad Then
      CurrentJournal "PJ", ES_SYSDATE, sJournalID
      
      FillCombo
      
      bOnLoad = False
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   'FormLoad Me, ES_DONTLIST, ES_DONTLIST
   FormLoad Me, ES_DONTLIST
   FormatControls
   sCurrForm = Caption
   bOnLoad = True
   
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set diaAPf01a = Nothing
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Private Sub FillCombo()
   'Dim RdoInv As ADODB.Recordset
   On Error GoTo DiaErr1
   
   FillVendors Me
   cmbVnd = cUR.CurrentVendor
   FindVendor Me
   GetInvoices
   
   Exit Sub
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub GetInvoices()
   Dim RdoInv As ADODB.Recordset
   Dim i As Long
   MouseCursor 13
   cmbInv.Clear
   On Error GoTo DiaErr1
   sSql = "SELECT VINO FROM VihdTable WHERE VIVENDOR = '" _
          & Compress(cmbVnd) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoInv)
   If bSqlRows Then
      cmbInv.enabled = True
      With RdoInv
         While Not .EOF
            AddComboStr cmbInv.hWnd, "" & Trim(.Fields(0))
            i = i + 1
            .MoveNext
         Wend
      End With
      cmbInv.ListIndex = 0
   Else
      cmbInv.enabled = False
   End If
   lblFnd = i
   Set RdoInv = Nothing
   MouseCursor 0
   Exit Sub
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Function CheckInvoice() As Byte
   Dim RdoChk As ADODB.Recordset
   Dim sInvoice As String
   Dim sVendor As String
   Dim sMsg As String
   
   On Error GoTo DiaErr1
   
   CheckInvoice = False
   sInvoice = Trim(cmbInv)
   sVendor = Compress(cmbVnd)
   
   ' 1. Check items for payment
   'sSql = "SELECT DISTINCT VITNO,VITVENDOR FROM " _
   '    & "ViitTable WHERE (VITNO='" & sInvoice & "' " _
   '    & "AND VITPAID=1 AND VITVENDOR='" _
   '    & sVendor & "')"
   
   sSql = "SELECT VIPAY FROM VihdTable WHERE VINO = '" & sInvoice _
          & "' AND VIVENDOR = '" & sVendor & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoChk)
   If bSqlRows Then
      With RdoChk
         If .Fields(0) > 0 Then
            sMsg = "Invoice Cancellation Not Allowed." & vbCrLf _
                   & "Invoice Has Payments Applied."
            MsgBox sMsg, vbInformation, Caption
            Exit Function
         End If
      End With
   End If
   Set RdoChk = Nothing
   
   ' 2. Check computer check setup
   sSql = "SELECT DISTINCT CHKNUM FROM ChseTable " _
          & "WHERE CHKINV = '" & sInvoice _
          & "' AND CHKVND = '" & sVendor & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoChk)
   If bSqlRows Then
      sMsg = "Invoice Cancellation Not Allowed." & vbCrLf _
             & "Payment For This Invoice Is Allocated" & vbCrLf _
             & "In Current Check Setup.  Clear Check" & vbCrLf _
             & "Setup And Then Cancel Invoice."
      MsgBox sMsg, vbInformation, Caption
      Exit Function
   End If
   Set RdoChk = Nothing
   
   ' 3. Check for closed journal
   sSql = "SELECT MJCLOSED,MJGLJRNL FROM JritTable INNER JOIN " _
          & "JrhdTable ON JritTable.DCHEAD = JrhdTable.MJGLJRNL " _
          & "WHERE DCVENDORINV = '" & sInvoice _
          & "' AND DCVENDOR = '" & sVendor _
          & "' AND MJTYPE = 'PJ'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoChk)
   If bSqlRows Then
      With RdoChk
         If Not IsNull(.Fields(0)) Then
            sMsg = "Invoice Cancellation Not Allowed." & vbCrLf _
                   & "Invoice Resides In Closed Journal " & .Fields(1) _
                   & "." & vbCrLf & "Reopen Journal Then Cancel Invoice."
            MsgBox sMsg, vbInformation, Caption
            Exit Function
         End If
      End With
   End If
   Set RdoChk = Nothing
   
   ' 4. Check For Closed MO
   sSql = "SELECT RUNCLOSED, RUNNO, PARTNUM FROM RunsTable INNER JOIN " _
          & "ViitTable ON RunsTable.RUNREF = ViitTable.VITMO AND " _
          & "RunsTable.RUNNO = ViitTable.VITMORUN INNER JOIN PartTable " _
          & "ON RunsTable.RUNREF = PartTable.PARTREF WHERE VITVENDOR = '" _
          & sVendor & "' AND VITNO = '" & sInvoice & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoChk)
   If bSqlRows Then
      With RdoChk
         If Not IsNull(.Fields(0)) Then
            sMsg = "Invoice Cancellation Not Allowed." & vbCrLf _
                   & "Invoice Cost Allocated To MO #" & vbCrLf _
                   & .Fields(1) & " Run # " & .Fields(2) & "."
            MsgBox sMsg, vbInformation, Caption
            Exit Function
         End If
      End With
   End If
   Set RdoChk = Nothing
   CheckInvoice = True
   Exit Function
   
DiaErr1:
   sProcName = "CheckInvoice"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Function

Public Sub CancelInvoice()
   Dim RdoInv As ADODB.Recordset
   'Dim rdoAct As ADODB.Recordset
   Dim bResponse As Byte
   Dim a As Integer
   Dim iTrans As Integer
   Dim iRef As Integer
   Dim sPon As String
   Dim sVendor As String
   Dim sInv As String
   
   On Error GoTo DiaErr1
   
   sMsg = "Are Certain That You Wish To " _
          & vbCrLf & "Cancel Invoice " & cmbInv & " for vendor " & cmbVnd & "?"
   bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
   If bResponse = vbYes Then
      
      'Err = 0
      'On Error Resume Next
      
      sVendor = Compress(cmbVnd)
      sInv = Trim(cmbInv)
      'If sJournalID <> "" Then iTrans = GetNextTransaction(sJournalID)
      
      clsADOCon.BeginTrans
      clsADOCon.ADOErrNum = 0
      
      sSql = "DELETE FROM JritTable WHERE (LEFT(DCHEAD,2) = 'PJ') AND (DCVENDOR = '" _
             & sVendor & "') AND (DCVENDORINV = '" & sInv & "')"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "SELECT VITNO,VITVENDOR,VITPO,VITPOITEM,VITPOITEMREV," _
             & "VIFREIGHT,VITAX FROM ViitTable,VihdTable WHERE " _
             & "(VITNO='" & Trim(cmbInv) & "' " _
             & "AND VITVENDOR='" & sVendor & "') AND VITNO=VINO"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoInv, ES_STATIC)
      If bSqlRows Then
         With RdoInv
            While Not .EOF
               If !VITPOITEM > 0 Then
                  'PO or just a DB or CM
                  sSql = "UPDATE PoitTable SET PITYPE=15 WHERE (PINUMBER=" _
                         & !VITPO & " AND PIITEM=" & !VITPOITEM & " AND PIREV='" _
                         & Trim(!VITPOITEMREV) & "') "
                  clsADOCon.ExecuteSQL sSql
                  
                  sPon = "PO " & Trim(!VITPO) & "- ITEM " & Trim(str(!VITPOITEM)) _
                         & Trim(!VITPOITEMREV)
                  
                  sSql = "UPDATE InvaTable SET INAMT=0 " _
                         & "WHERE INTYPE=15 AND INREF2='" & sPon & "' "
                  clsADOCon.ExecuteSQL sSql
               End If
               .MoveNext
            Wend
         End With
      End If
      Set RdoInv = Nothing
      
'      sSql = "DELETE FROM VihdTable WHERE (VINO='" & sInv & "' " _
'             & "AND VIVENDOR='" & sVendor & "') "
'      clsAdoCon.ExecuteSQL sSQL
'
      sSql = "DELETE FROM ViitTable WHERE (VITNO='" & sInv & "' " _
             & "AND VITVENDOR='" & sVendor & "') "
      clsADOCon.ExecuteSQL sSql
      
      sSql = "DELETE FROM SpapTable WHERE COAPINV = '" & sInv _
             & "' AND COAPVENDOR = '" & sVendor & "'"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "DELETE FROM VihdTable WHERE (VINO='" & sInv & "' " _
             & "AND VIVENDOR='" & sVendor & "') "
      clsADOCon.ExecuteSQL sSql
      
      If clsADOCon.ADOErrNum = 0 Then
         clsADOCon.CommitTrans
         SysMsg "Invoice Canceled.", True
         GetInvoices
         If cmbInv.enabled Then
            cmbInv.SetFocus
         End If
      Else
         clsADOCon.RollbackTrans
         clsADOCon.ADOErrNum = 0
         sMsg = "Could Not Cancel Invoice " & cmbInv & vbCrLf _
                & "Transaction Canceled."
         MsgBox sMsg, vbExclamation, Caption
      End If
   Else
      CancelTrans
   End If
   Exit Sub
   
DiaErr1:
   clsADOCon.RollbackTrans
   sProcName = "CancelInvoice"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub
