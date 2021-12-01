VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaARe01a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Customer Invoice (Sales Order)"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5670
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   5670
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox optCan 
      Caption         =   "OptCan"
      Height          =   255
      Left            =   480
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox cmbPre 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1920
      TabIndex        =   1
      ToolTipText     =   "Invoice Prefix(A-Z)"
      Top             =   1680
      Width           =   510
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Enabled         =   0   'False
      Height          =   315
      Left            =   4680
      TabIndex        =   11
      ToolTipText     =   "Add This Sales Order Invoice"
      Top             =   480
      Width           =   875
   End
   Begin VB.TextBox txtInv 
      Height          =   315
      Left            =   2410
      TabIndex        =   2
      ToolTipText     =   "Requires A Number"
      Top             =   1680
      Width           =   870
   End
   Begin VB.ComboBox cmbSon 
      Height          =   315
      Left            =   2160
      TabIndex        =   0
      Tag             =   "1"
      ToolTipText     =   "Contains Sales Orders With Items Not Shipped"
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   4680
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
      Autosize        =   2
      RoundedCorners  =   0   'False
      BevelWidth      =   0
      Outline         =   0   'False
      PictureUp       =   "diaARe01a.frx":0000
      PictureDn       =   "diaARe01a.frx":0146
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   3960
      Top             =   240
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2385
      FormDesignWidth =   5670
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Next Invoice Number"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   10
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label lblNme 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1920
      TabIndex        =   9
      Top             =   1320
      Width           =   2775
   End
   Begin VB.Label lblCst 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1920
      TabIndex        =   8
      Top             =   960
      Width           =   1395
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   7
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label lblPre 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1920
      TabIndex        =   6
      Top             =   600
      Width           =   255
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Order Number"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   615
      Width           =   1695
   End
End
Attribute VB_Name = "diaARe01a"
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

'***********************************************************************************
' diaARe01a - Post Sales Order Customer Invoice
'
' Notes:
'   - AR Invoice Types
'       CM = Credit Memo
'       DM = Debit Memo
'       CO = Credit Memo for Invoice
'       CR = Service Charge Invoice
'       SO = Sales Order (No Packslip)
'       PS = Sales Order (With Packslip)
'       CA = Advanced payment invoice
'       TM = Temporary (Place holder to lock record)
'       MS = Misc cash recipt
'
' Created: (cjs)
' Revisions:
'   11/17/01 (nth) Fixed error with invoice items falling behind the menu
'   04/01/02 (nth) Fixed error with closing form when sales journal is not open
'   08/26/02 (nth) Check for bCancel before running FormClose to prevent menu flash
'   12/18/02 (nth) Added check for taxable sales orders who's customer has no tax code.
'   01/08/02 (nth) Remember last invoice prefix, requested by Jevco
'   06/11/03 (nth) Removed the prompt to cancel the TM invoice
'
'***********************************************************************************

Dim RdoQry As ADODB.Command
Dim AdoParameter1 As ADODB.Parameter

Dim bOnLoad As Byte
Dim bGoodSO As Boolean
Dim bCancel As Boolean
Dim bInvUsed As Boolean
Dim bGoodJrn As Byte
Dim lNewInv As Long
Dim sInvPre As String
Dim sSoStadr As String
Dim sSOCust As String
Dim sMsg As String

Dim taxPerLineItem As Boolean


Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'***********************************************************************************

Private Sub cmbPre_LostFocus()
   cmbPre = CheckLen(cmbPre, 1)
   If Asc(cmbPre) < 65 Or Asc(cmbPre) > 90 Then
      Beep
      cmbPre = sInvPre
   End If
End Sub

Private Sub cmbSon_Click()
   bGoodSO = GetSalesOrder()
End Sub

Private Sub cmbSon_LostFocus()
   cmbSon = CheckLen(cmbSon, SO_NUM_SIZE)
   On Error Resume Next
   cmbSon = Format(Abs(Val(cmbSon)), SO_NUM_FORMAT)
   If bCancel Then
      Exit Sub
   Else
      bGoodSO = GetSalesOrder()
   End If
End Sub

Private Sub cmdAdd_Click()
   If Val(txtInv) <> lNewInv Then
      bInvUsed = GetOldInvoice(txtInv)
   Else
      bInvUsed = False
   End If
   
   If bInvUsed Then
      MsgBox "Invoice Number Is In Use.", vbInformation, Caption
      txtInv = Format(lNewInv, "000000")
   Else
      If bGoodSO Then
         UpdateInvoice
      Else
         MsgBox "Please Select A Valid Sales Order.", _
            vbInformation, Caption
      End If
   End If
   
End Sub

Private Sub cmdCan_Click()
   bCancel = True
   Unload Me
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   bCancel = True
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      SelectHelpTopic Me, "Customer Invoice (Sales Order)"
      cmdHlp = False
      MouseCursor 0
   End If
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      ' is tax calculated per line item?
      taxPerLineItem = IsTaxCalculatedPerLineItem()
      
      CurrentJournal "SJ", ES_SYSDATE, sJournalID
      bGoodJrn = GetSJAccounts()
      bGoodJrn = 1
      GetNextInvoice
      AddInvoice
      FillCombo
      GetOptions
      MouseCursor 0
      bOnLoad = False
   End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Debug.Print KeyCode
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   sCurrForm = Caption
   sInvPre = GetSetting("Esi2000", "EsiFina", "LastInvPref", sInvPre)
   If Len(Trim(sInvPre)) = 0 Then
      sInvPre = "I"
   End If
   On Error Resume Next
   sSql = "SELECT DISTINCT SONUMBER,SOTYPE,SOCUST,SOSTNAME,SOSTADR,ITSO," _
          & "CUREF,CUNICKNAME,CUNAME FROM SohdTable,SoitTable,CustTable " _
          & "WHERE SONUMBER=ITSO AND (" _
          & "ITQTY<>0 AND ITPSNUMBER='' AND ITINVOICE=0) AND (SOCUST=CUREF AND " _
          & "SONUMBER= ? )"
   Set RdoQry = New ADODB.Command
   RdoQry.CommandText = sSql
   
   Set AdoParameter1 = New ADODB.Parameter
   AdoParameter1.Type = adInteger
   RdoQry.parameters.Append AdoParameter1
   
   bOnLoad = True
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   Debug.Print "form mousedown"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   'Dim bResponse As Byte
   'Dim sMsg      As String
   
   'If bCancel Then
   '    sMsg = "Do You Really Want To Cancel The Addition" & vbCrLf _
   '        & "Of Invoice " & Trim(cmbPre) & txtInv & " ?"
   '    bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
   '
   '    If bResponse = vbYes Then
   '        Cancel = False
   '    Else
   '        Cancel = True
   '    End If
   'End If
   
   
   On Error Resume Next
   'dump any left over dummies
   sSql = "DELETE FROM CihdTable WHERE INVNO=" & lNewInv _
          & "AND INVTYPE='TM'"
   clsADOCon.ExecuteSQL sSql
   SaveSetting "Esi2000", "EsiFina", "LastInvPref", cmbPre
   SaveOptions
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   If bCancel Then FormUnload
   Set AdoParameter1 = Nothing
   
   Set RdoQry = Nothing
   Set diaARe01a = Nothing
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Public Sub FillCombo()
   Dim b As Byte
   Dim i As Integer
   Dim RdoSon As ADODB.Recordset
   cmbPre = sInvPre
   For i = 65 To 90
      AddComboStr cmbPre.hWnd, Chr$(i)
   Next
   
   On Error GoTo DiaErr1
   sProcName = "fillcombo"
   sSql = "SELECT DISTINCT SONUMBER,SOTYPE,ITSO FROM SohdTable,SoitTable " _
          & "WHERE SONUMBER=ITSO AND ITQTY<>0 AND ITPSNUMBER = '' AND ITINVOICE=0 " _
          & "AND ITCANCELED = 0 " _
          & "ORDER BY SONUMBER"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoSon)
   If bSqlRows Then
      With RdoSon
         cmbSon = Format(!SONUMBER, SO_NUM_FORMAT)
         Do Until .EOF
            AddComboStr cmbSon.hWnd, Format$(!SONUMBER, SO_NUM_FORMAT)
            .MoveNext
         Loop
         .Cancel
      End With
   End If
   If cmbSon.ListCount > 0 Then bGoodSO = GetSalesOrder()
   Set RdoSon = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Public Function GetSalesOrder() As Boolean
   Dim RdoSon As ADODB.Recordset
   
   On Error GoTo DiaErr1
   RdoQry.parameters(0).Value = Val(cmbSon)
   bSqlRows = clsADOCon.GetQuerySet(RdoSon, RdoQry)
   If bSqlRows Then
      With RdoSon
         lblCst.ForeColor = vbBlack
         lblNme.ForeColor = vbBlack
         cmbSon = Format(!SONUMBER, SO_NUM_FORMAT)
         lblPre = "" & Trim(!SOTYPE)
         lblCst = "" & Trim(!CUNICKNAME)
         lblNme = "" & Trim(!CUNAME)
         sSOCust = "" & Trim(!CUREF)
         sSoStadr = "" & Trim(!SOSTNAME) & vbCrLf _
                    & Trim(!SOSTADR)
         .Cancel
         If bGoodJrn Then cmdAdd.enabled = True
         GetSalesOrder = True
      End With
   Else
      cmdAdd.enabled = False
      GetSalesOrder = False
      lblPre = ""
      lblCst.ForeColor = ES_RED
      lblNme.ForeColor = ES_RED
      lblCst = "** Invalid **"
      lblNme = "No Such Sales Order Or No Items To Ship."
   End If
   Set RdoSon = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getsalesor"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Function

Public Sub GetNextInvoice()
'   Dim RdoInv As ADODB.RecordSet
'
'   On Error GoTo DiaErr1
'   sSql = "SELECT MAX(INVNO) FROM CihdTable WHERE INVNO<999999"
'   bSqlRows = clsAdoCon.GetDataSet(sSql,RdoInv)
'   If bSqlRows Then
'      If IsNull(RdoInv.Fields(0)) Then
'         txtInv = "000001"
'      Else
'         lNewInv = (RdoInv.Fields(0) + 1)
'         txtInv = Format(lNewInv, "000000")
'      End If
'   Else
'      txtInv = "000001"
'   End If
'   Set RdoInv = Nothing
   Dim inv As New ClassARInvoice
   lNewInv = inv.GetNextInvoiceNumber
   txtInv = Format(lNewInv, "000000")
   Exit Sub
   
DiaErr1:
   sProcName = "getnextinv"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub txtInv_LostFocus()
   txtInv = CheckLen(txtInv, 6)
   txtInv = Format(Abs(Val(txtInv)), "000000")
End Sub

Public Sub AddInvoice()
   'Reserve a record in case the invoice number is changed.
   'Use TM so that it won't show and can be safely deleted.
   On Error GoTo DiaErr1
   sSql = "INSERT INTO CihdTable (INVNO,INVTYPE,INVSO,INVUSR) " _
          & "VALUES(" & lNewInv & ",'TM'," _
          & Val(cmbSon) & ",'" & Secure.UserInitials & "')"
   clsADOCon.ExecuteSQL sSql
   Exit Sub

   
DiaErr1:
   sProcName = "addinvoice"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Public Sub UpdateInvoice()
   Dim rdoTSo As ADODB.Recordset
   
   On Error GoTo DiaErr1
   
   sSql = "SELECT CUTAXCODE, SOTAXABLE " _
          & "FROM CustTable INNER JOIN " _
          & "SohdTable ON CustTable.CUREF = SohdTable.SOCUST " _
          & "WHERE SONUMBER = " & Val(cmbSon)
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoTSo)
   
   If rdoTSo!SOTAXABLE = 1 And ("" & Trim(rdoTSo!CUTAXCODE)) = "" _
                         And Not taxPerLineItem Then
      sMsg = "Sales Order " & lblPre & cmbSon & " Is Taxable" & vbCrLf _
             & "But " & lblCst & " Has No Tax Code Assigned."
      MsgBox sMsg, vbInformation, Caption
   Else
      lCurrInvoice = Val(txtInv)
      
      sSql = "UPDATE CihdTable SET INVNO=" & lCurrInvoice & "," _
             & "INVPRE='" & cmbPre & "',INVTYPE='SO'," _
             & "INVSTADR='" & Left(sSoStadr, 255) & "'," _
             & "INVSO=" & Val(cmbSon) & ",INVCUST='" & sSOCust & "' " _
             & "WHERE INVNO=" & lNewInv & " AND INVTYPE='TM'"
      clsADOCon.ExecuteSQL sSql
      
      Dim inv As New ClassARInvoice
      inv.SaveLastInvoiceNumber lCurrInvoice
      
      diaARe01b.lblSon = lblPre & cmbSon
      diaARe01b.lblInv = cmbPre & txtInv
      diaARe01b.lblCst = lblCst
      diaARe01b.lblNme = lblNme
      diaARe01b.txtStAdr = sSoStadr
      diaARe01b.optLoad = vbChecked
      
      
      
      Unload Me
      diaARe01b.Show
   End If
   
   Set rdoTSo = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "updateinvo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

' See if the Accounts are there

Public Function GetSJAccounts() As Byte
   Dim rdoJrn As ADODB.Recordset
   Dim i As Integer
   Dim b As Byte
   
   On Error GoTo DiaErr1
   sSql = "SELECT COREF,COSJARACCT,COSJINVACCT,COSJNFRTACCT," _
          & "COSJTFRTACCT,COSJTAXACCT FROM ComnTable WHERE " _
          & "COREF=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoJrn, ES_FORWARD)
   If bSqlRows Then
      With rdoJrn
         For i = 1 To 5
            If Not IsNull(.Fields(i)) Then
               If Trim(.Fields(i)) = "" Then b = 1
            Else
               b = 1
            End If
         Next
         .Cancel
      End With
   End If
   If b = 1 Then
      MsgBox "One or more missing journal account(s): " _
         & "COSJARACCT, COSJINVACCT, COSJNFRTACCT, COSJTFRTACCT, COSJTAXACCT", _
         vbInformation, Caption
      GetSJAccounts = 0
   Else
      GetSJAccounts = 1
   End If
   Set rdoJrn = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getsjacco"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Function

Public Sub SaveOptions()
   Dim sOptions As String
   sOptions = Trim(cmbPre)
   SaveSetting "Esi2000", "EsiFina", Me.Name, sOptions
End Sub

Public Sub GetOptions()
   Dim sOptions As String
   sOptions = Trim(GetSetting("Esi2000", "EsiFina", Me.Name, sOptions))
   If sOptions <> "" Then cmbPre = sOptions
   
End Sub

Private Function IsTaxCalculatedPerLineItem() As Boolean
   Dim rdo As ADODB.Recordset
   Dim result As Boolean
   
   result = False
   
   On Error GoTo done
   sSql = "SELECT TaxPerItem from ComnTable"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdo, ES_FORWARD)
   If bSqlRows Then
      result = rdo.Fields(0)
   End If
   Set rdo = Nothing
done:
   IsTaxCalculatedPerLineItem = result
End Function
