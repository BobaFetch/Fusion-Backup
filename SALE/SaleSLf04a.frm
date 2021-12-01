VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form SaleSLf04a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change A Customer Nickname"
   ClientHeight    =   2820
   ClientLeft      =   3000
   ClientTop       =   1710
   ClientWidth     =   6255
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   6255
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "SaleSLf04a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.TextBox txtNew 
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      Tag             =   "3"
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton cmdDel 
      Cancel          =   -1  'True
      Caption         =   "&Apply"
      Enabled         =   0   'False
      Height          =   315
      Left            =   5280
      TabIndex        =   6
      ToolTipText     =   "Delete The Current Customer"
      Top             =   600
      Width           =   915
   End
   Begin VB.ComboBox cmbCst 
      Height          =   315
      Left            =   1680
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "3"
      Top             =   1080
      Width           =   1555
   End
   Begin VB.TextBox txtNme 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Tag             =   "2"
      Top             =   1440
      Width           =   3475
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Close"
      Height          =   435
      Left            =   5280
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   915
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5880
      Top             =   2400
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2820
      FormDesignWidth =   6255
   End
   Begin MSComctlLib.ProgressBar prg1 
      Height          =   300
      Left            =   1440
      TabIndex        =   12
      Top             =   2400
      Width           =   4212
      _ExtentX        =   7435
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label lblWrn 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "A Customer Is Recorded With That Nickname"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   1680
      TabIndex        =   10
      Top             =   2160
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "New Nickname"
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   9
      Top             =   1800
      Width           =   1425
   End
   Begin VB.Label lblWrn 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   8
      Top             =   360
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.Label lblWrn 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Please Close All Other Sections Before Proceeding"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   7
      Top             =   600
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Current Nickname"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   1095
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Name"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   4
      Top             =   1440
      Width           =   1425
   End
End
Attribute VB_Name = "SaleSLf04a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'5/26/06 Added CpayTable, ArecTable
'6/20/06 Added Illegal Character Trap
Option Explicit
Dim bOnLoad As Byte
Dim sNewCust As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtNme.BackColor = BackColor
   
End Sub

Private Function CheckWindows() As Byte
   Dim b As Byte
   b = Val(GetSetting("Esi2000", "Sections", "admn", 0))
   b = b + Val(GetSetting("Esi2000", "Sections", "prod", 0))
   b = b + Val(GetSetting("Esi2000", "Sections", "engr", 0))
   b = b + Val(GetSetting("Esi2000", "Sections", "fina", 0))
   b = b + Val(GetSetting("Esi2000", "Sections", "qual", 0))
   If b > 0 Then
      lblWrn(0) = sSysCaption & " Has Determined " & b & " Other Open Section(s)"
      lblWrn(0).Visible = True
      lblWrn(1).Visible = True
      cmdDel.Enabled = False
   End If
   CheckWindows = b
   
End Function

Private Sub cmbCst_Click()
   GetDelCustomer
   
End Sub


Private Sub cmbCst_LostFocus()
   cmbCst = CheckLen(cmbCst, 10)
   If Len(Trim(cmbCst)) Then GetDelCustomer
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub



Private Sub cmdDel_Click()
   Dim b As Byte
   If Trim(txtNew) = "" Then Exit Sub
   b = TestNewCustomer()
   If b = 0 Then
      MsgBox "That Customer Has Been Previously Installed.", _
         vbInformation, Caption
      Exit Sub
   End If
   If txtNme.ForeColor = ES_RED Then
      MsgBox "Requires A Valid Customer.", _
         vbInformation, Caption
   Else
      ChangeTheCustomer
   End If
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 2154
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      CheckWindows
      FillCombo
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST
   FormatControls
   
   lblWrn(0).ForeColor = ES_RED
   lblWrn(1).ForeColor = ES_RED
   lblWrn(2).ForeColor = ES_RED
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set SaleSLf04a = Nothing
   
End Sub



Private Sub FillCombo()
   On Error GoTo DiaErr1
   MouseCursor 13
   cmbCst.Clear
   sSql = "SELECT CUREF,CUNICKNAME FROM CustTable "
   LoadComboBox cmbCst
   MouseCursor 0
   If cmbCst.ListCount > 0 Then cmbCst = cmbCst.List(0)
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub



Private Sub GetDelCustomer()
   Dim RdoCst As ADODB.Recordset
   On Error GoTo DiaErr1
   MouseCursor 13
   sSql = "Qry_GetCustomerBasics '" & Compress(cmbCst) & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCst, ES_FORWARD)
   If bSqlRows Then
      With RdoCst
         cmbCst = "" & Trim(!CUNICKNAME)
         txtNme = "" & Trim(!CUNAME)
         ClearResultSet RdoCst
      End With
   Else
      txtNme = "*** Customer Wasn't Found ***"
   End If
   Set RdoCst = Nothing
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "getdelcust"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub



Private Sub txtNew_Change()
   lblWrn(2).Visible = False
   
End Sub

Private Sub txtNew_LostFocus()
   txtNew = CheckLen(txtNew, 10)
   
End Sub


Private Sub txtNme_Change()
   If Left(txtNme, 6) = "*** Cu" Then
      txtNme.ForeColor = ES_RED
      cmdDel.Enabled = False
   Else
      txtNme.ForeColor = Es_TextForeColor
      cmdDel.Enabled = True
   End If
   
End Sub



Private Sub ChangeTheCustomer()
   Dim bResponse As Byte
   Dim sMsg As String
   Dim sCust As String
   
   bResponse = IllegalCharacters(sNewCust)
   If bResponse > 0 Then
      MsgBox "The New Nickname Contains An Illegal " & Chr$(bResponse) & ".", _
         vbInformation, Caption
      Exit Sub
   End If
   
   sMsg = "It Is Not A Good Idea To Change A Customer's Nickname " & vbCrLf _
          & "If There Is Any Chance That It Is In Use Right Now."
   MsgBox sMsg, vbExclamation, Caption
   
   sMsg = "This Function Permanently Changes The Customer " & vbCrLf _
          & "Nickname Are You Sure That You Want To Continue?      "
   bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
   sCust = Compress(cmbCst)
   If bResponse = vbYes Then
      On Error Resume Next
      MouseCursor 13
      prg1.Visible = True
      clsADOCon.BeginTrans
      clsADOCon.ADOErrNum = 0
      'start checking
      'Sales Orders
      
      sSql = "INSERT INTO CustTable (CUREF,CUNICKNAME,CUNUMBER,CUNAME,CUADR,CUBTNAME,CUBTADR,CUSEL," & vbCrLf
      sSql = sSql & " CUCOL,CUTERMS,CUTELEPHONE,CUDISCOUNT,CUCREDIT,CUARPCT,CUTYPE,CUTAXCODE,CUTAXSTATE," & vbCrLf
      sSql = sSql & " CUFAX,CUDIVISION,CUREGION,CUSALESMAN,CUBALANCE,CUCUTOFF,CUARBALANCE,CUARDISC,CUDAYS," & vbCrLf
      sSql = sSql & " CUNETDAYS,CUSVCACCT,CUTEMP,CUDATE,CUFOB,CUSTERMS,CUVIA,CUSTNAME,CUSTADR,CUSLSMAN,CUBORTAXCODE," & vbCrLf
      sSql = sSql & " CUBOWTAXCODE,CUFRTALLOW,CUFRTDAYS,CUCCONTACT,CUCPHONE,CUCEXT,CUTAXEXEMPT,CUDPP,CUCREATEDBY," & vbCrLf
      sSql = sSql & " CUCREATEDDT,CUREVISED,CUREP,CUCRDEXP,CUDBRATEDT,CUDBRATE,CUAPCONTACT,CUAPPHONE,CUAPPHONEEXT," & vbCrLf
      sSql = sSql & " CUAPFAX,CUVEND,CUMASTER,CUWEB,CUEMAIL,CUSTATE,CUBSTATE,CUSTSTATE,CUCITY,CUBCITY,CUSTCITY,CUZIP," & vbCrLf
      sSql = sSql & " CUBZIP,CUSTZIP,CUAREACODE,CUQAPHONE,CUQAPHONEEXT,CUQAFAX,CUQAEMAIL,CUQAREP,CUCOUNTRY,CUBCOUNTRY," & vbCrLf
      sSql = sSql & " CUSTCOUNTRY,CUPRICEBOOK,CUQBNAME,CUBOTAXACCT,CUSALESACCT,CUINTPHONE,CUINTFAX,CUCINTPHONE,CUAPINTPHONE," & vbCrLf
      sSql = sSql & " CUAPINTFAX,CUQAINTPHONE,CUQAINTFAX,CUSINCE,CUBEOM,CUALLOWTRANSFERS,CUEINVOICING,CUAPEMAIL,CUCREDITLIMIT," & vbCrLf
      sSql = sSql & " CUPRINTKANBAN,CUPRINTPACCAR,CUSTBLDG,CUSTDOOR)" & vbCrLf
      sSql = sSql & "SELECT '" & sNewCust & "','" & sNewCust & "',CUNUMBER,CUNAME,CUADR,CUBTNAME,CUBTADR,CUSEL," & vbCrLf
      sSql = sSql & " CUCOL,CUTERMS,CUTELEPHONE,CUDISCOUNT,CUCREDIT,CUARPCT,CUTYPE,CUTAXCODE,CUTAXSTATE," & vbCrLf
      sSql = sSql & " CUFAX,CUDIVISION,CUREGION,CUSALESMAN,CUBALANCE,CUCUTOFF,CUARBALANCE,CUARDISC,CUDAYS," & vbCrLf
      sSql = sSql & " CUNETDAYS,CUSVCACCT,CUTEMP,CUDATE,CUFOB,CUSTERMS,CUVIA,CUSTNAME,CUSTADR,CUSLSMAN,CUBORTAXCODE," & vbCrLf
      sSql = sSql & " CUBOWTAXCODE,CUFRTALLOW,CUFRTDAYS,CUCCONTACT,CUCPHONE,CUCEXT,CUTAXEXEMPT,CUDPP,CUCREATEDBY," & vbCrLf
      sSql = sSql & " CUCREATEDDT,CUREVISED,CUREP,CUCRDEXP,CUDBRATEDT,CUDBRATE,CUAPCONTACT,CUAPPHONE,CUAPPHONEEXT," & vbCrLf
      sSql = sSql & " CUAPFAX,CUVEND,CUMASTER,CUWEB,CUEMAIL,CUSTATE,CUBSTATE,CUSTSTATE,CUCITY,CUBCITY,CUSTCITY,CUZIP," & vbCrLf
      sSql = sSql & " CUBZIP,CUSTZIP,CUAREACODE,CUQAPHONE,CUQAPHONEEXT,CUQAFAX,CUQAEMAIL,CUQAREP,CUCOUNTRY,CUBCOUNTRY," & vbCrLf
      sSql = sSql & " CUSTCOUNTRY,CUPRICEBOOK,CUQBNAME,CUBOTAXACCT,CUSALESACCT,CUINTPHONE,CUINTFAX,CUCINTPHONE,CUAPINTPHONE," & vbCrLf
      sSql = sSql & " CUAPINTFAX,CUQAINTPHONE,CUQAINTFAX,CUSINCE,CUBEOM,CUALLOWTRANSFERS,CUEINVOICING,CUAPEMAIL,CUCREDITLIMIT," & vbCrLf
      sSql = sSql & " CUPRINTKANBAN,CUPRINTPACCAR,CUSTBLDG,CUSTDOOR FROM CustTable" & vbCrLf
      sSql = sSql & "WHERE CUREF = '" & sCust & "'"
      
        Debug.Print sSql
        
      clsADOCon.ExecuteSQL sSql 'rdExecDirect

      sSql = "UPDATE SohdTable SET SOCUST='" & sNewCust & "' " _
             & "WHERE SOCUST='" & sCust & "'"
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
      prg1.Value = 10
      
      'Shouldn't be any, but test PS anyway
      sSql = "UPDATE PshdTable SET PSCUST='" & sNewCust & "' " _
             & "WHERE PSCUST='" & sCust & "'"
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
      prg1.Value = 20
      
      'Invoice?
      sSql = "UPDATE CihdTable SET INVCUST='" & sNewCust & "' " _
             & "WHERE INVCUST='" & sCust & "'"
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
      prg1.Value = 30
      
      'Journal?
      sSql = "UPDATE JritTable SET DCCUST='" & sNewCust & "' " _
             & "WHERE DCCUST='" & sCust & "'"
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
      prg1.Value = 40
      
      'RejTag?
      sSql = "UPDATE RjhdTable SET REJCUST='" & sNewCust & "' " _
             & "WHERE REJCUST='" & sCust & "'"
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
      prg1.Value = 50
      
      'Document?
      sSql = "UPDATE DdocTable SET DOCUST='" & sNewCust & "' " _
             & "WHERE DOCUST='" & sCust & "'"
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
      prg1.Value = 60
      
      'Cash
      sSql = "UPDATE CashTable SET CACUST='" & sNewCust & "' " _
             & "WHERE CACUST='" & sCust & "'"
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
      prg1.Value = 70
      
      'Estimating?
      sSql = "UPDATE EstiTable SET BIDCUST='" & sNewCust & "' " _
             & "WHERE BIDCUST='" & sCust & "'"
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
      prg1.Value = 80
      
      'RFQ's?
      sSql = "UPDATE RfqsTable SET RFQCUST='" & sNewCust & "' " _
             & "WHERE RFQCUST='" & sCust & "'"
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
      prg1.Value = 85
      
      'Tools
      sSql = "UPDATE TohdTable SET TOOL_CUST='" & sNewCust & "' " _
             & "WHERE TOOL_CUST='" & sCust & "'"
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
      prg1.Value = 90
      
      sSql = "UPDATE LoitTable SET LOICUST='" & sNewCust & "' " _
             & "WHERE LOICUST='" & sCust & "'"
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
      
      '8/23/05
      sSql = "UPDATE LohdTable SET LOTCUST='" & sNewCust & "' " _
             & "WHERE LOTCUST='" & sCust & "'"
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
      prg1.Value = 95
      
      '5/26/06
      sSql = "UPDATE CpayTable SET CPCUST='" & sNewCust & "' " _
             & "WHERE CPCUST='" & sCust & "'"
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
      
      '5/26/06
      sSql = "UPDATE ArecTable SET RECCUST='" & sNewCust & "' " _
             & "WHERE RECCUST='" & sCust & "'"
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
      
      prg1.Value = 95
      
      MouseCursor 0
      sMsg = "Last Chance. Are You Sure That You Want" & vbCrLf _
             & "To Change Customer " & cmbCst & "?"
      bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
      If bResponse = vbYes Then
         'customer - let's do it
         sSql = "UPDATE CustTable SET CUNICKNAME='" & Trim(txtNew) & "' WHERE " _
                & "CUREF='" & sNewCust & "'"
         clsADOCon.ExecuteSQL sSql 'rdExecDirect
         ' Delete the old customer
         sSql = "DELETE FROM CustTable WHERE CUREF='" & sCust & "'"
         
         clsADOCon.ExecuteSQL sSql 'rdExecDirect
         prg1.Value = 100
         If clsADOCon.ADOErrNum = 0 Then
            clsADOCon.CommitTrans
            SysMsg "Nickname Was Changed.", True
            cUR.CurrentCustomer = cmbCst
            txtNew = ""
            FillCombo
         Else
            clsADOCon.RollbackTrans
            MsgBox "Could Not Change The Customer Nickname.", _
               vbExclamation, Caption
         End If
      Else
         clsADOCon.RollbackTrans
         CancelTrans
      End If
   Else
      CancelTrans
   End If
   prg1.Visible = False
   
End Sub

Private Function TestNewCustomer() As Byte
   Dim RdoTst As ADODB.Recordset
   sNewCust = Compress(txtNew)
   If sNewCust = "" Then
      TestNewCustomer = 0
      Exit Function
   End If
   On Error GoTo DiaErr1
   MouseCursor 13
   sSql = "SELECT CUREF FROM CustTable WHERE CUREF='" & sNewCust & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoTst, ES_FORWARD)
   If bSqlRows Then
      ClearResultSet RdoTst
      lblWrn(2).Visible = True
      TestNewCustomer = 0
   Else
      lblWrn(2).Visible = False
      TestNewCustomer = 1
   End If
   Set RdoTst = Nothing
   MouseCursor 0
   Exit Function
   
DiaErr1:
   sProcName = "testnewcu"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function
