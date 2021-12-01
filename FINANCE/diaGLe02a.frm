VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form diaGLe02a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Journal Entry"
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8670
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   8670
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdRevTran 
      Caption         =   "Reverse &Trans"
      Enabled         =   0   'False
      Height          =   315
      Left            =   6120
      TabIndex        =   41
      ToolTipText     =   "Reverse Journal Entry"
      Top             =   2280
      Width           =   1350
   End
   Begin VB.CommandButton cmdPrintJournal 
      Height          =   375
      Left            =   8040
      Picture         =   "diaGLe02a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   40
      ToolTipText     =   "Print Journal Entry to Printer"
      Top             =   1320
      Width           =   495
   End
   Begin VB.TextBox txtExt 
      Height          =   975
      Left            =   1080
      MultiLine       =   -1  'True
      TabIndex        =   2
      Tag             =   "9"
      Top             =   1080
      Width           =   3375
   End
   Begin VB.CheckBox chkTem 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   7680
      TabIndex        =   4
      Top             =   1800
      Width           =   855
   End
   Begin VB.CommandButton cmdCnl 
      Caption         =   "&Reselect"
      Enabled         =   0   'False
      Height          =   315
      Left            =   7680
      TabIndex        =   35
      ToolTipText     =   "Select Another Journal"
      Top             =   2280
      Width           =   875
   End
   Begin VB.CommandButton cmdSel 
      Caption         =   "&Select"
      Enabled         =   0   'False
      Height          =   315
      Left            =   7680
      TabIndex        =   5
      ToolTipText     =   "Display GL Journal Items"
      Top             =   600
      Width           =   875
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "&Delete"
      Enabled         =   0   'False
      Height          =   315
      Left            =   7680
      TabIndex        =   27
      Top             =   6480
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   7200
      Top             =   120
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   6885
      FormDesignWidth =   8670
   End
   Begin VB.TextBox txtDesc 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Tag             =   "2"
      Top             =   720
      Width           =   2775
   End
   Begin VB.ComboBox txtPst 
      Height          =   315
      Left            =   5640
      TabIndex        =   3
      Tag             =   "4"
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox txtCmt 
      Height          =   285
      Left            =   4800
      TabIndex        =   12
      Tag             =   "2"
      Top             =   6120
      Width           =   2775
   End
   Begin VB.TextBox txtCrd 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3480
      TabIndex        =   11
      Tag             =   "1"
      Top             =   6120
      Width           =   1200
   End
   Begin VB.TextBox txtDeb 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2160
      TabIndex        =   10
      Tag             =   "1"
      Top             =   6120
      Width           =   1200
   End
   Begin VB.ComboBox cmbAct 
      Height          =   315
      Left            =   720
      TabIndex        =   9
      Tag             =   "3"
      Top             =   6120
      Width           =   1335
   End
   Begin VB.TextBox txtRef 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   6120
      Width           =   495
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      Height          =   315
      Left            =   6720
      TabIndex        =   13
      Top             =   6480
      Width           =   875
   End
   Begin VB.CommandButton cmdPst 
      Caption         =   "&Post"
      Height          =   315
      Left            =   7680
      TabIndex        =   6
      ToolTipText     =   "Post GL Journal"
      Top             =   960
      Width           =   875
   End
   Begin VB.ComboBox cmbTran 
      Height          =   315
      Left            =   1320
      TabIndex        =   7
      Tag             =   "1"
      Top             =   2280
      Width           =   735
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   3135
      Left            =   120
      TabIndex        =   8
      Top             =   2880
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   5530
      _Version        =   393216
      Rows            =   1
      Cols            =   5
      FixedRows       =   0
      FixedCols       =   0
      BackColor       =   16777215
      ScrollBars      =   2
      SelectionMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ComboBox cmbJrn 
      Height          =   315
      Left            =   1080
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "2"
      ToolTipText     =   "Select From List"
      Top             =   360
      Width           =   1530
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   7680
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   15
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
      PictureUp       =   "diaGLe02a.frx":044B
      PictureDn       =   "diaGLe02a.frx":0591
   End
   Begin Threed.SSFrame z2 
      Height          =   135
      Index           =   0
      Left            =   120
      TabIndex        =   16
      Top             =   2040
      Width           =   8445
      _Version        =   65536
      _ExtentX        =   14896
      _ExtentY        =   238
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblRevJrlName 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3720
      TabIndex        =   42
      Top             =   6480
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Extended Description"
      Height          =   375
      Index           =   13
      Left            =   120
      TabIndex        =   39
      Top             =   960
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Template"
      Height          =   255
      Index           =   12
      Left            =   6840
      TabIndex        =   38
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label lblCurCrd 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3480
      TabIndex        =   37
      Top             =   2280
      Width           =   1200
   End
   Begin VB.Label lblCurDeb 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2160
      TabIndex        =   36
      Top             =   2280
      Width           =   1200
   End
   Begin VB.Label lblDif 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5640
      TabIndex        =   34
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label lblDeb 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5640
      TabIndex        =   33
      ToolTipText     =   "Sum Of All Debits For This Journal"
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label lblCrd 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5640
      TabIndex        =   32
      ToolTipText     =   "Sum Of All Credits For This Journal"
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label lblActDesc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   720
      TabIndex        =   31
      Top             =   6480
      Width           =   2775
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   4680
      X2              =   6840
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Difference"
      Height          =   255
      Index           =   10
      Left            =   4680
      TabIndex        =   30
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Credit"
      Height          =   255
      Index           =   9
      Left            =   4680
      TabIndex        =   29
      Top             =   720
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Debit"
      Height          =   255
      Index           =   1
      Left            =   4680
      TabIndex        =   28
      Top             =   360
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   26
      Top             =   720
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Post Date"
      Height          =   255
      Index           =   11
      Left            =   4680
      TabIndex        =   24
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ref     "
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
      Index           =   8
      Left            =   120
      TabIndex        =   23
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Transaction"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   22
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Credit Amt           "
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
      Index           =   6
      Left            =   3480
      TabIndex        =   21
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Comments                                                                              "
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
      Index           =   5
      Left            =   4800
      TabIndex        =   20
      Top             =   2640
      Width           =   3735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Debit Amt            "
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
      Left            =   2160
      TabIndex        =   19
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Account               "
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
      Left            =   840
      TabIndex        =   18
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Journal ID"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   17
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "diaGLe02a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of                     ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***

'See the UpdateTables prodecure for database revisions

'************************************************************************************
'
' diaGLe02a - Add/Revise GL Journal Entries
'
' Created: (nth)
' Revsions:
' 02/14/03 (nth) Increased numerical data sizes and added CURRENCYMASK per JLH
' 05/13/03 (nth) Added KEYSET cursor lookups and updates.
' 01/01/04 (JCW) Fixed invalid result set update when updating \ Prevent editing blank Jrnl
' 04/14/04 (nth) Added template option.
' 06/08/04 (nth) Added extended description per JLH and THYPRE.
' 09/21/04 (nth) Added trap for invalide GL account.
' 01/19/05 (nth) Added trap for posting date without fiscal period.
'
'************************************************************************************

Option Explicit

Dim bOnLoad As Byte
Dim bCancel As Byte
Dim bGoodId As Byte
Dim bGoodYear As Byte
Dim bGoodJrn As Byte ' Good Journal yes of no
Dim bDisTrn As Byte ' Disable the click and lostfocus events for cmbtran

Dim rdoJrn As ADODB.Recordset ' Current Journal Header
Dim RdoItm As ADODB.Recordset ' Current GL Journal Items
Dim sJournal As String

' 1 = AddTran
' 2 = ReviseTran
' 0 = Nothing
Dim bTransType As Byte

Dim iFyear As Integer
Dim iJrnNo As Integer
Dim sJrnl As String
Dim sMsg As String

' Numeric mask used in this transaction formats up to 1 billion $.
Const ROWSPERSCREEN = 13 ' Show 13 grid rows

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'************************************************************************************

Private Sub sManageBoxs(Optional bItemsMode As Byte)
   If bItemsMode Then
      cmbJrn.enabled = False
      txtDesc.enabled = False
      txtPst.enabled = False
      cmdSel.enabled = False
      Grid1.enabled = True
      cmbAct.enabled = True
      cmbTran.enabled = True
      txtDeb.enabled = True
      txtcrd.enabled = True
      txtCmt.enabled = True
      cmdUpdate.enabled = True
      cmdDel.enabled = True
      cmdCnl.enabled = True
      lblActDesc.enabled = True
   Else
      cmbJrn.enabled = True
      txtDesc.enabled = True
      txtPst.enabled = True
      cmdSel.enabled = True
      Grid1.enabled = False
      Grid1.Clear
      Grid1.Rows = 0
      cmbAct.enabled = False
      cmbTran.enabled = False
      cmbTran.Clear
      cmbTran = ""
      lblCurDeb = ""
      lblCurCrd = ""
      txtDeb.enabled = False
      txtcrd.enabled = False
      txtCmt.enabled = False
      cmdUpdate.enabled = False
      cmdDel.enabled = False
      cmdCnl.enabled = False
      lblActDesc.enabled = False
   End If
End Sub


Public Sub sPostGLJournal()
   Dim bResponse As Byte
   
   On Error GoTo DiaErr1
   
   sMsg = "Do You Wish To Post General Journal " _
          & Trim(cmbJrn) & " ?"
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   
   If bResponse = vbYes Then
      Err = 0
      On Error Resume Next
      clsADOCon.BeginTrans
      clsADOCon.ADOErrNum = 0
      
      sSql = "UPDATE GjhdTable SET " _
             & "GJPOSTED = 1" _
             & " WHERE GJNAME = '" & sJournal & "'"
      clsADOCon.ExecuteSQL sSql
      
      If clsADOCon.ADOErrNum = 0 Then
         clsADOCon.CommitTrans
         SysMsg sJournal & " Successfully Posted.", 1
         sManageBoxs
         sFillCombo
         cmbJrn.SetFocus
         bGoodJrn = False
      Else
         clsADOCon.RollbackTrans
         clsADOCon.ADOErrNum = 0
         MouseCursor 0
         MsgBox "Could Not Post " & sJournal & ".", _
            vbExclamation, Caption
      End If
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "sPostGLJournal"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Function fMaxRef() As Integer
   Dim rdoRef As ADODB.Recordset
   
   On Error GoTo DiaErr1
   
   ' Get next reference number
   sSql = "SELECT Max(JIREF) AS MaxOfJIREF FROM GjitTable " _
          & "WHERE JINAME = '" & Trim(cmbJrn) & " ' AND JITRAN = " _
          & CInt(cmbTran)
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoRef, ES_FORWARD)
   
   With rdoRef
      If IsNull(!MaxOfJIREF) Then
         fMaxRef = 0
      Else
         fMaxRef = !MaxOfJIREF
      End If
   End With
   Set rdoRef = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "sfMaxRef"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Function

Private Function fOpenGLJournal() As Byte
   Dim rdoItems As ADODB.Recordset
   On Error GoTo DiaErr1
   ' Get Journal Header
   sSql = "SELECT GJPOST,GJDESC,GJTEMPLATE,GJEXTDESC FROM GjhdTable " _
          & "WHERE GJNAME = '" & sJournal & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoJrn, ES_KEYSET)
   If bSqlRows Then
      With rdoJrn
         txtDesc = "" & Trim(!GJDESC)
         txtPst = Format(!GJPOST, DATEMASK)
         chkTem = Val("" & !GJTEMPLATE)
         txtExt = "" & Trim(!GJEXTDESC)
      End With
      ' Flag set to sum all transactions is GL journal
      sSumCurrentTran True
      fOpenGLJournal = True
   End If
   Set rdoItems = Nothing
   Exit Function
DiaErr1:
   sProcName = "fOpenGLJournal"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Function

Private Sub sCreateGLJournal()
   Dim iResponse As Integer
   'Dim rdoYr As ADODB.Recordset
   Dim i As Integer
   Dim dNow As Date
   Dim sPst As String
   
   On Error GoTo DiaErr1
   
   sMsg = "Journal " & sJournal & _
          " Does Not Exists.  Do You Wish To Create It?"
   iResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   
   ' Add GL journal to database
   If iResponse = vbYes Then
      
'      dNow = Format(GetServerDateTime, "mm/dd/yy")
'
'      sSql = "SELECT FYPERSTART1,FYPEREND1,FYPERSTART2,FYPEREND2,FYPERSTART3,FYPEREND3," _
'             & "FYPERSTART4,FYPEREND4,FYPERSTART5,FYPEREND5,FYPERSTART6,FYPEREND6,FYPERSTART7," _
'             & "FYPEREND7,FYPERSTART8,FYPEREND8,FYPERSTART9,FYPEREND9,FYPERSTART10,FYPEREND10," _
'             & "FYPERSTART11,FYPEREND11,FYPERSTART12,FYPEREND12,FYPERSTART13,FYPEREND13 " _
'             & "From GlfyTable Where FYYEAR = " & Format(dNow, "yyyy")
'      bSqlRows = clsAdoCon.GetDataSet(sSql,rdoYr)
'
'      If bSqlRows Then
'         With rdoYr
'            For i = 0 To 25 Step 2
'               If dNow >= .Fields(i) And dNow <= .Fields(i + 1) Then
'                  sPst = Format(.Fields(i + 1), "mm/dd/yy")
'               End If
'            Next
'            .Cancel
'         End With
'         Set rdoYr = Nothing
'      End If
      
      sPst = GetFYPeriodEnd(GetServerDateTime)
      
      On Error Resume Next
      clsADOCon.BeginTrans
      clsADOCon.ADOErrNum = 0
      
      sSql = "INSERT INTO GjhdTable (GJNAME,GJOPEN,GJPOST) " _
             & " VALUES ('" & sJournal & "','" & dNow & "','" & sPst & "')"
      clsADOCon.ExecuteSQL sSql
      
      If clsADOCon.ADOErrNum = 0 Then
         clsADOCon.CommitTrans
         sMsg = sJournal & " Successfully Created."
         SysMsg sMsg, 1, Me
         bGoodJrn = fOpenGLJournal()
      Else
         clsADOCon.RollbackTrans
         clsADOCon.ADOErrNum = 0
         MsgBox "Could Not Create " & sJournal & ".", _
            vbExclamation, Caption
      End If
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "sCreateGLJournal"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub chkTem_Click()
   With rdoJrn
      !GJTEMPLATE = chkTem.Value
      .Update
   End With
End Sub

Private Sub cmbAct_Click()
   lblActDesc = UpdateActDesc(cmbAct)
End Sub

Private Sub cmbAct_LostFocus()
   ' No multiple account selections allowed here
   If Trim(UCase(cmbAct)) = "ALL" _
           Or Trim(cmbAct) = "" Then
      lblActDesc.ForeColor = ES_RED
      lblActDesc = "*** Invalid Account Number ***"
   Else
      lblActDesc = UpdateActDesc(cmbAct)
   End If
End Sub

Private Sub cmbTran_Click()
   If Not bDisTrn Then
      sFillGLItems
   End If
End Sub

Private Sub cmbTran_LostFocus()
   If Not bDisTrn Then
      sFillGLItems
   End If
End Sub

Private Sub cmdCnl_Click()
   sManageBoxs
   Set RdoItm = Nothing
   cmbJrn.SetFocus
End Sub

Private Sub cmdCnl_MouseDown(Button As Integer, _
                             Shift As Integer, X As Single, Y As Single)
   bDisTrn = True
End Sub

Private Sub cmdCnl_MouseUp(Button As Integer, Shift As Integer, _
                           X As Single, Y As Single)
   bDisTrn = False
End Sub

Private Sub cmdDel_Click()
   sDeleteGLItem
End Sub

Private Sub cmdPrintJournal_Click()
    diaGLp03a.bRemote = 1
    diaGLp03a.Visible = False
    diaGLp03a.cmbJrn = Me.cmbJrn
    diaGLp03a.optPrn.Value = True
    'diaGLp03a.PrintReport1 (Me.cmbJrn)
End Sub

Private Sub cmdPst_Click()
   sPostGLJournal
End Sub

Private Function ReverseGlJrn(strSrcGj As String, iTrans As Integer, strNewGj As String) As Boolean
   Dim rdoJrn1 As ADODB.Recordset
   Dim rdoJrn2 As ADODB.Recordset
   Dim sMsg As String
   
   On Error GoTo DiaErr1
   MouseCursor 13
   
   ' Copy Header
   Err.Clear
   On Error Resume Next
   clsADOCon.BeginTrans
   clsADOCon.ADOErrNum = 0
   
   sSql = "SELECT * FROM GjhdTable WHERE GJNAME = '" & strSrcGj & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoJrn1)
   
   With rdoJrn1
      sSql = "INSERT INTO GjhdTable (GJNAME,GJDESC,GJPOST,GJOPEN,GJPOSTED) " _
             & "VALUES (" _
             & "'" & strNewGj & "'," _
             & "'" & !GJDESC & "'," _
             & "'" & !GJPOST & "'," _
             & "'" & !GJOPEN & "'," _
             & "0)"
      
      Debug.Print sSql
      clsADOCon.ExecuteSQL sSql
   End With
   Set rdoJrn1 = Nothing
   
   ' Copy items
   sSql = "SELECT * FROM GjitTable WHERE JINAME = '" & strSrcGj & "' AND JITRAN = " & iTrans
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoJrn2, ES_KEYSET)
   
   With rdoJrn2
      While Not .EOF
         sSql = "INSERT INTO GjitTable (JINAME,JIDESC,JITRAN,JIREF," _
                & "JIACCOUNT,JIDEB,JICRD) " _
                & "VALUES (" _
                & "'" & strNewGj & "'," _
                & "'" & Trim(!JIDESC) & "'," _
                & !JITRAN & "," _
                & !JIREF & "," _
                & "'" & Trim(!JIACCOUNT) & "'" _
                & "," & !JICRD & "," & !JIDEB & ")"
         
         Debug.Print sSql
         
         clsADOCon.ExecuteSQL sSql
         
         .MoveNext
      Wend
   End With
   Set rdoJrn2 = Nothing
   
   If clsADOCon.ADOErrNum = 0 Then
      clsADOCon.CommitTrans
      sMsg = "Reversed Debit/Credit Transactions from " & strSrcGj & " To " & strNewGj & " ."
      MsgBox sMsg, vbInformation, Caption
      ReverseGlJrn = True
   Else
      clsADOCon.RollbackTrans
      clsADOCon.ADOErrNum = 0
      sMsg = "Could Not Copy " & strSrcGj & " To " & strNewGj & " ."
      MsgBox sMsg, vbInformation, Caption
      ReverseGlJrn = False
   End If
   MouseCursor 0
   Exit Function
   
DiaErr1:
   sProcName = "CopyGlJrn"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Function

Private Sub cmdRevTran_Click()

Dim strSrcGj As String
Dim iTrans As Integer
Dim strNewGj As String

strSrcGj = cmbJrn.Text
iTrans = IIf(cmbTran <> "", Val(cmbTran), 0)

diaNewJournal.Show 1
strNewGj = lblRevJrlName

If (strNewGj <> "") Then
   
   Dim bRet As Boolean
   bRet = ReverseGlJrn(strSrcGj, iTrans, strNewGj)
   
   If (bRet = True) Then
      cmbJrn.AddItem (strNewGj)
   End If
Else
   sMsg = "Select New GL Journal Name."
   MsgBox sMsg, vbInformation, Caption
End If

End Sub

Private Sub cmdSel_Click()
   If Trim(cmbJrn) <> "" Then
      sManageBoxs True
      sFillGLTrans
      sFillGLItems
   
      ' if GL Posted, don't allow editing the journal
      Dim bPosted As Boolean
      bPosted = CheckGLPosted(sJournal)
      If (bPosted = True) Then
         cmdUpdate.enabled = False
      End If
   
      cmdRevTran.enabled = True
      
   Else
      MsgBox "Enter A Journal ID.", vbInformation, Caption
      cmbJrn.SetFocus
   End If
End Sub

Private Sub Grid1_Click()
   sGetGridRow
End Sub

Private Sub Grid1_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = 46 Then
      sDeleteGLItem
   ElseIf KeyCode = 13 Then
      sGetGridRow
   End If
End Sub

Private Sub Grid1_LostFocus()
   If Not bDisTrn Then
      sGetGridRow
   End If
End Sub

Private Sub Grid1_MouseDown(Button As Integer, _
                            Shift As Integer, X As Single, Y As Single)
   bDisTrn = True
End Sub

Private Sub Grid1_MouseUp(Button As Integer, Shift As Integer, _
                          X As Single, Y As Single)
   bDisTrn = False
End Sub

Private Sub txtCmt_LostFocus()
   txtCmt = CheckLen(txtCmt, 30)
   txtCmt = CheckComments(txtCmt)
   txtCmt = StrCase(txtCmt)
End Sub

Private Sub sCalcTotals()
   sSumCurrentTran
   sSumCurrentTran True 'all tranactions
End Sub

Private Sub sGetGridRow()
   On Error Resume Next
   ' Grab the info on the row
   Grid1.Row = Grid1.RowSel
   Grid1.Col = 0
   
   ' Did we click on the add new entry row?
   If Trim(Grid1) = "*" Then
      bTransType = 1
      txtRef = fMaxRef + 1
      txtDeb = "0.00"
      txtcrd = "0.00"
      cmdDel.enabled = False
      cmbAct.SetFocus
   Else
      bTransType = 2
      txtRef = Grid1
      Grid1.Col = 1
      cmbAct.Text = Trim(Grid1)
      Grid1.Col = 2
      txtDeb = Grid1
      Grid1.Col = 3
      txtcrd = Grid1
      Grid1.Col = 4
      txtCmt = Trim(Grid1)
      cmdDel.enabled = True
   End If
   
   ' if GL Posted, don't allow editing the journal
   Dim bPosted As Boolean
   bPosted = CheckGLPosted(sJournal)
   If (bPosted = True) Then
      cmdDel.enabled = False
   End If
   
End Sub

Public Function fCheckFiscalYear() As Byte
   Dim RdoFyr As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT FYYEAR FROM GlfyTable "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoFyr, ES_FORWARD)
   If bSqlRows Then
      fCheckFiscalYear = 1
      RdoFyr.Cancel
   Else
      fCheckFiscalYear = 0
   End If
   Set RdoFyr = Nothing
   Exit Function
DiaErr1:
   sProcName = "checkfisc"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Function

Private Sub cmbjrn_Click()
   If Not bCancel Then
      sJournal = Trim(cmbJrn)
      bGoodJrn = fOpenGLJournal
   End If
End Sub

Private Sub cmbjrn_LostFocus()
   Dim bPosted As Boolean
   If Not bCancel Then
      cmbJrn = CheckLen(cmbJrn, 12)
      sJournal = Trim(cmbJrn)
      If Trim(cmbJrn) <> "" Then
         bGoodJrn = fOpenGLJournal
         If Not bGoodJrn Then
            sCreateGLJournal
         End If
         
         bPosted = CheckGLPosted(sJournal)
         
         If (bPosted = True) Then
            cmdPst.enabled = False
            cmdUpdate.enabled = False
            cmdDel.enabled = False
         End If
         
      End If
   End If
End Sub

Private Function CheckGLPosted(ByVal strJournal As String) As Boolean
   Dim rdoGL As ADODB.Recordset
   
   sSql = "SELECT DISTINCT GJPOSTED FROM GjhdTable WHERE GJNAME = '" & strJournal & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoGL, ES_FORWARD)
   If bSqlRows Then
      With rdoGL
         If Not rdoGL.EOF Then
            CheckGLPosted = IIf(!GJPOSTED = 0, False, True)
         End If
      End With
   Else
      CheckGLPosted = False
   End If
   Set rdoGL = Nothing

End Function
Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, _
                             X As Single, Y As Single)
   bCancel = True
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      SelectHelpTopic Me, "Journal Entry"
      cmdHlp = False
      MouseCursor 0
   End If
End Sub

Private Sub cmdUpdate_Click()
   sUpdateGLItem
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      bGoodYear = fCheckFiscalYear()
      If bGoodYear Then
         sFillCombo
      Else
         Dim bResponse As Byte
         sMsg = "Fiscal Years Have Not Been Initialized." & vbCr _
                & "Initialize Fiscal Years Now?"
         bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
         If bResponse = vbYes Then
            Unload Me
            diaGLe04a.Show
         Else
            Unload Me
         End If
      End If
      bOnLoad = False
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   sCurrForm = Caption
   With Grid1
      .Rows = 0
      .Cols = 5
      .ColWidth(0) = 500
      .ColWidth(1) = 1500
      .ColWidth(2) = 1300
      .ColWidth(3) = 1300
      .ColWidth(4) = (.Width - 5000)
   End With
   sManageBoxs
   bOnLoad = True
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set rdoJrn = Nothing
   Set RdoItm = Nothing
   Set diaGLe02a = Nothing
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Public Sub sFillCombo()
   Dim rdoJrn As ADODB.Recordset
   Dim rdoAct As ADODB.Recordset
   
   On Error GoTo DiaErr1
   
   cmbJrn.Clear
   sSql = "SELECT GJNAME FROM GjhdTable WHERE GJPOSTED = 0"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoJrn, ES_FORWARD)
   If bSqlRows Then
      With rdoJrn
         While Not .EOF
            AddComboStr cmbJrn.hWnd, "" & Trim(!GJNAME)
            .MoveNext
         Wend
      End With
      rdoJrn.Cancel
      cmbJrn.ListIndex = 0 ' default to first one
   Else
      cmdPst.enabled = False
      txtDesc = ""
      txtRef = ""
   End If
   Set rdoJrn = Nothing
   
   
   ' Fill account combo / need to add account descriptions
   cmbAct.Clear
   sSql = "Qry_FillLowAccounts"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoAct, ES_FORWARD)
   If bSqlRows Then
      With rdoAct
         Do Until .EOF
            AddComboStr cmbAct.hWnd, "" & Trim(!GLACCTNO)
            .MoveNext
         Loop
      End With
      rdoAct.Cancel
      cmbAct.ListIndex = 0
   End If
   Set rdoAct = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "sFillCombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub txtCrd_LostFocus()
   If Trim(txtcrd) = "" Then txtcrd = "0.00"
   txtcrd = Format(txtcrd, CURRENCYMASK)
End Sub

Private Sub txtDeb_LostFocus()
   If Trim(txtDeb) = "" Then txtDeb = "0.00"
   txtDeb = Format(txtDeb, CURRENCYMASK)
End Sub

Private Sub txtDesc_GotFocus()
   SelectFormat Me
End Sub

Private Sub txtDesc_LostFocus()
   txtDesc = CheckLen(txtDesc, 30)
   txtDesc = StrCase(txtDesc)
   txtDesc = CheckComments(txtDesc)
   If bGoodJrn Then
      On Error Resume Next
      With rdoJrn
         !GJDESC = "" & Trim(txtDesc)
         .Update
      End With
      If Err > 0 Then
         ValidateEdit Me
      End If
   End If
End Sub

Private Sub txtExt_GotFocus()
   SelectFormat Me
End Sub

Private Sub txtExt_LostFocus()
   txtExt = CheckLen(txtExt, 512)
   txtExt = CheckComments(txtExt)
   If bGoodJrn Then
      On Error Resume Next
      With rdoJrn
         !GJEXTDESC = Trim(txtExt)
         .Update
         If Err > 0 Then
            ValidateEdit Me
         End If
      End With
   End If
End Sub

Private Sub txtPst_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtPst_GotFocus()
   SelectFormat Me
End Sub

Private Sub txtPst_LostFocus()
   txtPst = CheckDate(txtPst)
   If GetFYPeriodEnd(txtPst) = "" Then
      sMsg = "No Fiscal Period Defined For Posting Date " & txtPst & "."
      MsgBox sMsg, vbInformation, Caption
      If Not bCancel Then
         txtPst.SetFocus
      End If
      Exit Sub
   End If
   If bGoodJrn Then
      On Error Resume Next
      With rdoJrn
         !GJPOST = txtPst
         .Update
      End With
      If Err > 0 Then
         ValidateEdit Me
      End If
   End If
End Sub

Private Sub sFillGLItems()
   Dim sEntry As String
   
   On Error GoTo DiaErr1
   
   Set RdoItm = Nothing
   MouseCursor 13
   
   sSql = "SELECT JINAME,JIDESC,JICRD,JIDEB,JITRAN,JIREF,JIACCOUNT,JILASTREVBY " _
          & "FROM GjitTable WHERE JINAME = '" & sJournal & "' AND " _
          & "JITRAN = " & Trim(cmbTran) & " ORDER BY JIREF"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoItm, ES_KEYSET)
   If bSqlRows Then
      Grid1.Clear
      Grid1.Rows = 0
      With RdoItm
         While Not .EOF
            sEntry = "" _
                     & CStr(!JIREF) & Chr(9) _
                     & " " & Trim(!JIACCOUNT) & Chr(9) _
                     & Format(!JIDEB, CURRENCYMASK) & Chr(9) _
                     & Format(!JICRD, CURRENCYMASK) & Chr(9) _
                     & " " & Trim(!JIDESC)
            Grid1.AddItem sEntry
            .MoveNext
         Wend
         .Cancel
      End With
   Else
      AddComboStr cmbTran.hWnd, cmbTran
      Grid1.Clear
      Grid1.Rows = 0
   End If
   
   sAddNewItemRow
   sCalcTotals
   
   'Grid1.Row = Grid1.RowSel
   Grid1.Row = Grid1.Rows - 1
   Grid1.RowSel = Grid1.Row
   Grid1.Col = 0
   sGetGridRow
   
   If Grid1.Rows > ROWSPERSCREEN Then
      Grid1.TopRow = Grid1.Rows - ROWSPERSCREEN
   End If
   
   
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "sFillGLItems"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub sFillGLTrans()
   Dim iLastTran As Integer
   Dim rdoTrn As ADODB.Recordset
   
   On Error GoTo DiaErr1
   
   bDisTrn = True
   
   If Trim(cmbTran) = "" Then
      iLastTran = 1
   Else
      iLastTran = CInt(cmbTran)
   End If
   
   cmbTran.Clear
   
   sSql = "SELECT DISTINCT JITRAN FROM GjitTable " _
          & "WHERE JINAME = '" & sJournal & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoTrn, ES_FORWARD)
   With rdoTrn
      While Not .EOF
         AddComboStr cmbTran.hWnd, "" & !JITRAN
         .MoveNext
      Wend
      rdoTrn.Cancel
   End With
   
   Set rdoTrn = Nothing
   
   cmbTran = iLastTran
   bDisTrn = False
   Exit Sub
   
DiaErr1:
   sProcName = "sFillGLTrans"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub sSumCurrentTran(Optional bAllTrans As Byte)
   Dim RdoSum As ADODB.Recordset
   
   On Error GoTo DiaErr1
   sSql = "SELECT SUM(JICRD),SUM(JIDEB) FROM GjitTable " _
          & "WHERE JINAME = '" & sJournal & "'"
   
   If bAllTrans Then
      ' Sum ALL transactions.
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoSum, ES_FORWARD)
      If bSqlRows Then
         With RdoSum
            lblCrd = Format(.Fields(0), CURRENCYMASK)
            lblDeb = Format(.Fields(1), CURRENCYMASK)
            lblDif = Format(Abs(.Fields(0) - .Fields(1)), _
                     CURRENCYMASK)
         End With
      End If
      If lblCrd <> "" And lblDeb <> "" Then
         If lblCrd = lblDeb Then
            cmdPst.enabled = True
         Else
            cmdPst.enabled = False
         End If
      Else
         cmdPst.enabled = False
      End If
   Else
      sSql = sSql & " AND JITRAN = '" & Trim(cmbTran) & "'"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoSum, ES_FORWARD)
      If bSqlRows Then
         With RdoSum
            lblCurCrd = Format(.Fields(0), CURRENCYMASK)
            lblCurDeb = Format(.Fields(1), CURRENCYMASK)
         End With
      End If
   End If
   RdoSum.Cancel
   Set RdoSum = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "sSumCurrentTran"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
   
End Sub

Private Sub sAddNewItemRow()
   Dim sEntry As String
   On Error Resume Next
   sEntry = "*" & Chr(9) & Chr(9) & Chr(9) & _
            Chr(9) & "(Add New Entry) *"
   Grid1.AddItem sEntry
   
End Sub

Private Sub sDeleteGLItem()
   Dim bUsed As Byte
   
   On Error Resume Next
   sSql = "DELETE FROM GjitTable WHERE " _
          & "JINAME = '" & sJournal & "' AND " _
          & "JITRAN = " & Val(cmbTran) & " AND " _
          & "JIREF  = " & Val(txtRef)
   clsADOCon.ExecuteSQL sSql
   bUsed = clsADOCon.RowsAffected
   If Err > 0 Then bUsed = 0
   If bUsed Then
      SysMsg "Journal Entry Item Deleted.", True
   Else
      sMsg = "Cannot Delete Journal Item." & vbCrLf _
             & "Journal May Still Be In Use."
      MsgBox sMsg, vbInformation, Caption
   End If
   sFillGLItems
End Sub

Private Sub sUpdateGLItem()
   
   On Error GoTo DiaErr1
   sMsg = ""
   'EDIT # 1 VAL(TXTCRd)
   If CCur(Val(txtDeb)) <> 0 And CCur(Val(txtcrd)) <> 0 Then
      sMsg = "A Journal Entry Cannot Have Both " & vbCrLf _
             & "Debit And Credit Amounts."
   ElseIf CCur(Val(txtDeb)) = 0 And CCur(Val(txtcrd)) = 0 Then
      sMsg = "A Journal Entry Must Have Either " & vbCrLf _
             & "A Credit Or Debit Amount."
   ElseIf Left(lblActDesc, 6) = "*** In" Then
      sMsg = "Invalide Account Number."
   End If
   
   If sMsg = "" Then
      On Error Resume Next
      With RdoItm
         Select Case bTransType
            Case 1
               .AddNew
               !JINAME = Trim(cmbJrn)
               !JITRAN = CInt(cmbTran)
               !JIREF = CInt(txtRef)
            Case 2
               '.MoveFirst
               .Move CLng(Grid1.Row), 1
            Case Else
               Exit Sub
         End Select
         
         !JIACCOUNT = Compress(cmbAct)
         !JIDEB = CCur(txtDeb)
         !JICRD = CCur(txtcrd)
         !JIDESC = "" & Trim(txtCmt)
         !JILASTREVBY = Secure.UserInitials
         .Update
      End With
      
      ' If database transactions are successfully then update grid...
      If Err > 0 Then
         ValidateEdit Me
      Else
         Grid1.Col = 0
         Grid1.Text = txtRef
         Grid1.Col = 1
         Grid1.Text = " " & cmbAct
         Grid1.Col = 2
         Grid1.Text = Format(txtDeb, CURRENCYMASK)
         Grid1.Col = 3
         Grid1.Text = Format(txtcrd, CURRENCYMASK)
         Grid1.Col = 4
         Grid1.Text = " " & Trim(txtCmt)
         
         If bTransType = 1 Then
            sAddNewItemRow
            Grid1.Row = Grid1.Row + 1
            If Grid1.Rows > ROWSPERSCREEN Then
               Grid1.TopRow = Grid1.TopRow + 1 'auto scroll
            End If
            txtRef = Val(txtRef) + 1
            cmbAct.SetFocus
         End If
         
         ' Clear out old values but keep account # and comments.
         txtDeb = "0.00"
         txtcrd = "0.00"
         sCalcTotals
      End If
   Else
      MsgBox sMsg, vbInformation, Caption
      txtDeb.SetFocus
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "sUpdateGLItem"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub
