VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form PurcPRe05a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Buyers"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5520
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "PurcPRe05a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   14
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CheckBox optVew 
      Caption         =   "View"
      Height          =   255
      Left            =   1440
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdVnd 
      Cancel          =   -1  'True
      Caption         =   "&Vendors"
      Enabled         =   0   'False
      Height          =   315
      Left            =   4560
      TabIndex        =   6
      ToolTipText     =   "Assign Vendors To This Buyer"
      Top             =   720
      Width           =   875
   End
   Begin VB.CommandButton cmdCde 
      Caption         =   "&Apply"
      Enabled         =   0   'False
      Height          =   315
      Left            =   4560
      TabIndex        =   5
      ToolTipText     =   "Update All Parts With The Selected Product Code To This Buyer (MRP Incuded)"
      Top             =   2280
      Width           =   875
   End
   Begin VB.ComboBox cmbCde 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1800
      Sorted          =   -1  'True
      TabIndex        =   4
      Tag             =   "8"
      ToolTipText     =   "Select Product Code From List"
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox txtMid 
      Height          =   285
      Left            =   1800
      TabIndex        =   2
      Tag             =   "3"
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox txtLst 
      Height          =   285
      Left            =   1800
      TabIndex        =   3
      Tag             =   "2"
      Top             =   1920
      Width           =   2295
   End
   Begin VB.TextBox txtFst 
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Tag             =   "2"
      Top             =   1200
      Width           =   2295
   End
   Begin VB.ComboBox cmbByr 
      Height          =   315
      Left            =   1800
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Enter Or Revise A Buyer (20 Char Max)"
      Top             =   720
      Width           =   2535
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Close"
      Height          =   435
      Left            =   4560
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6000
      Top             =   3120
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3195
      FormDesignWidth =   5520
   End
   Begin MSComctlLib.ProgressBar prg1 
      Height          =   300
      Left            =   1800
      TabIndex        =   15
      Top             =   2760
      Width           =   3372
      _ExtentX        =   5953
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Code(s)"
      Height          =   255
      Index           =   4
      Left            =   360
      TabIndex        =   12
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Middle Initial"
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   11
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Last Name"
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   10
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "First Name"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   9
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Buyer ID"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   8
      Top             =   720
      Width           =   1335
   End
End
Attribute VB_Name = "PurcPRe05a"
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
Dim RdoBuy As ADODB.Recordset
Dim bCanceled As Boolean
Dim bOnLoad As Byte
Dim bGoodBuyer As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbByr_Change()
    If Len(cmbByr) > 20 Then cmbByr = Left(cmbByr, 20)
    ' Get the product code for the buyer
    GetProductCode (Trim(cmbByr))
   
End Sub

Private Sub cmbByr_Click()
   bGoodBuyer = GetThisBuyer
   
End Sub


Private Sub cmbByr_LostFocus()
   cmbByr = CheckLen(cmbByr, 20)
   If Not bCanceled Then
      bGoodBuyer = GetThisBuyer()
      If Len(Trim(cmbByr)) > 0 Then
         If bGoodBuyer = 0 Then AddBuyer
        
        ' Get the product code for the buyer
        GetProductCode (Trim(cmbByr))
      
      Else
         txtFst = ""
         txtLst = ""
         txtMid = ""
      End If
   End If
   
End Sub


Private Sub cmbCde_LostFocus()
   cmbCde = CheckLen(cmbCde, 6)
   ' 02/02/2009 This is the bug
   'If cmbCde.ListCount > 0 Then cmbCde = cmbCde.List(0)
   
End Sub


Private Sub cmdCan_Click()
   bCanceled = True
   Unload Me
   
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   bCanceled = True
   
End Sub


Private Sub cmdCde_Click()
   If bGoodBuyer = 1 Then UpdateProductCodes
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 4305
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub


Private Sub cmdVnd_Click()
   optVew.Value = vbChecked
   PurcPRe05b.lblByr = cmbByr
   PurcPRe05b.Show
   
End Sub


Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      FillBuyers
      FillProductCodes
      If cmbCde.ListCount > 0 Then cmbCde = cmbCde.List(0)
      If cmbByr.ListCount > 0 Then
         cmbByr = cmbByr.List(0)
         bGoodBuyer = GetThisBuyer
      End If
      bOnLoad = 0
   End If
   If optVew.Value = vbChecked Then
      optVew.Value = vbUnchecked
      Unload PurcPRe05b
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set RdoBuy = Nothing
   Set PurcPRe05a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub


Private Sub txtFst_LostFocus()
   txtFst = CheckLen(txtFst, 20)
   txtFst = StrCase(txtFst)
   If bGoodBuyer = 1 Then
      RdoBuy!BYFSTNAME = "" & txtFst
      RdoBuy.Update
   End If
   
End Sub


Private Sub txtLst_LostFocus()
   txtLst = CheckLen(txtLst, 20)
   txtLst = StrCase(txtLst)
   If bGoodBuyer = 1 Then
      RdoBuy!BYLSTNAME = "" & txtLst
      RdoBuy.Update
   End If
   
End Sub


Private Sub txtMid_LostFocus()
   txtMid = CheckLen(txtMid, 1)
   If bGoodBuyer = 1 Then
      RdoBuy!BYMIDINIT = "" & txtMid
      RdoBuy.Update
   End If
   
   
End Sub


Private Function GetThisBuyer() As Byte
   sSql = "SELECT * FROM BuyrTable WHERE BYREF='" _
          & Compress(cmbByr) & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoBuy, ES_KEYSET)
   If bSqlRows Then
      With RdoBuy
         txtLst = "" & Trim(!BYLSTNAME)
         txtFst = "" & Trim(!BYFSTNAME)
         txtMid = "" & Trim(!BYMIDINIT)
         ClearResultSet RdoBuy
         If cmbCde.ListCount > 0 Then cmdCde.Enabled = True
         cmdVnd.Enabled = True
         GetThisBuyer = 1
      End With
   Else
      cmdCde.Enabled = False
      cmdVnd.Enabled = False
      txtFst = ""
      txtLst = ""
      txtMid = ""
      GetThisBuyer = 0
   End If
   Exit Function
   
DiaErr1:
   sProcName = "getthisbu"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub AddBuyer()
   Dim bResponse As Byte
   Dim sBuyer As String
   Dim sMsg As String
   
   sMsg = cmbByr & " Was Not Found. Add This Buyer?"
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      On Error Resume Next
      clsADOCon.ADOErrNum = 0
      
      sBuyer = Compress(cmbByr)
      If sBuyer = "ALL" Then
         MsgBox "ALL Is A Reverved Word And Cannot " & vbCr _
            & "Be Used.  Please Select Another Number.", _
            vbInformation, Caption
         cmbByr = ""
         cmbByr.SetFocus
      Else
         sSql = "INSERT INTO BuyrTable (BYREF,BYNUMBER) " _
                & "VALUES('" & sBuyer & "','" & cmbByr & "')"
         clsADOCon.ExecuteSQL sSql
         If clsADOCon.ADOErrNum = 0 Then
            SysMsg "The Buyer Has Been Added.", True
            cmbByr.AddItem cmbByr
            bGoodBuyer = GetThisBuyer()
            cmbByr.SetFocus
         Else
            MsgBox "Couldn't Establish That Buyer.", _
               vbInformation, Caption
            cmbByr.SetFocus
         End If
      End If
   Else
      CancelTrans
   End If
   
End Sub

Private Sub UpdateProductCodes()
   Dim bResponse As Byte
   Dim sBuyer As String
   Dim sCode As String
   Dim sMsg As String
   
   sMsg = "Caution, This Operation Overwrites Any Existing" & vbCr _
          & "Settings With The Current Buyer. Are " & vbCr _
          & "You Sure That ou Wish To Continue?"
   bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
   If bResponse = vbYes Then
      MouseCursor 13
      cmdCan.Enabled = False
      cmdCde.Enabled = False
      On Error Resume Next
      sBuyer = Compress(cmbByr)
      sCode = Compress(cmbCde)
      prg1.Visible = True
      prg1.Value = 5
      clsADOCon.BeginTrans
      clsADOCon.ADOErrNum = 0
      
      On Error GoTo 0
      ' Update to the MRPL table is not necessary as data should be comming
      ' from existing PO table.
      'MM sSql = "UPDATE MrplTable SET MRP_POBUYER='" & sBuyer & "' " _
      'MM        & "WHERE MRP_PARTPRODCODE='" & sCode & "'"
      'MM clsAdoCon.ExecuteSQL sSql
      
      prg1.Value = 35
      sSql = "UPDATE PartTable SET PABUYER='" & sBuyer & "' " _
             & "WHERE PAPRODCODE='" & sCode & "'"
      clsADOCon.ExecuteSQL sSql
      
      prg1.Value = 70
      sSql = "UPDATE PcodTable SET PCBUYERREF='" & sBuyer & "' " _
             & "WHERE PCREF='" & sCode & "'"
      clsADOCon.ExecuteSQL sSql
      If clsADOCon.ADOErrNum = 0 Then
         prg1.Value = 80
         clsADOCon.CommitTrans
         Sleep 500
         sSql = "DELETE FROM BuycTable WHERE " _
                & "BYPRODCODE='" & sCode & "'"
         clsADOCon.ExecuteSQL sSql
         
         sSql = "INSERT INTO BuycTable (BYREF,BYPRODCODE) " _
                & "VALUES('" & sBuyer & "','" & sCode & "')"
         clsADOCon.ExecuteSQL sSql
         prg1.Value = 100
         MouseCursor 0
         MsgBox "Update Was Successful.", _
            vbInformation, Caption
      Else
         clsADOCon.RollbackTrans
         MouseCursor 0
         MsgBox "Update Was Not Successful.", _
            vbExclamation, Caption
      End If
      prg1.Visible = False
      cmdCan.Enabled = True
      cmdCde.Enabled = True
   Else
      CancelTrans
   End If
End Sub

Private Sub GetProductCode(sBuyer As String)
    
    On Error GoTo DiaErr1
    If Len(sBuyer) > 0 Then
        
        Dim RdoBuyer As ADODB.Recordset
        
        sSql = "SELECT BYPRODCODE FROM BuycTable WHERE BYREF='" _
            & sBuyer & "' "
        bSqlRows = clsADOCon.GetDataSet(sSql, RdoBuyer, ES_KEYSET)
        If bSqlRows Then
            With RdoBuyer
               cmbCde = "" & Trim(!BYPRODCODE)
               ClearResultSet RdoBuyer
            End With
        Else
            If cmbCde.ListCount > 0 Then cmbCde = cmbCde.List(0)
        End If
        
        Exit Sub
        
DiaErr1:
        sProcName = "GetProductCode"
        CurrError.Number = Err.Number
        CurrError.Description = Err.Description
        DoModuleErrors Me
        
    End If
End Sub


