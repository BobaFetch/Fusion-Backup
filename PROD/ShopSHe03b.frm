VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form ShopSHe03b 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Operation Completions"
   ClientHeight    =   8115
   ClientLeft      =   1740
   ClientTop       =   1065
   ClientWidth     =   10680
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8115
   ScaleWidth      =   10680
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCompleteAll 
      Cancel          =   -1  'True
      Caption         =   "Complete All"
      Height          =   480
      Left            =   9660
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   480
      Width           =   875
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "ShopSHe03b.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   35
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   250
   End
   Begin VB.TextBox lblCurList 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   1140
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   33
      TabStop         =   0   'False
      Tag             =   "2"
      Text            =   " "
      ToolTipText     =   "Click To View The Tool List"
      Top             =   6540
      Width           =   3075
   End
   Begin VB.TextBox lblLst 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   1140
      Locked          =   -1  'True
      TabIndex        =   32
      TabStop         =   0   'False
      Tag             =   "2"
      Text            =   " "
      Top             =   6900
      Width           =   3075
   End
   Begin VB.CommandButton cmdLst 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   9255
      Picture         =   "ShopSHe03b.frx":07AE
      Style           =   1  'Graphical
      TabIndex        =   31
      TabStop         =   0   'False
      ToolTipText     =   "Previous Entry"
      Top             =   6720
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.CommandButton cmdNxt 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   9780
      Picture         =   "ShopSHe03b.frx":08E4
      Style           =   1  'Graphical
      TabIndex        =   30
      TabStop         =   0   'False
      ToolTipText     =   "Next Entry"
      Top             =   6720
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.CheckBox optCom 
      Alignment       =   1  'Right Justify
      Caption         =   "Complete"
      Height          =   255
      Left            =   5460
      TabIndex        =   5
      ToolTipText     =   "Mark Operation Complete"
      Top             =   3660
      Width           =   1095
   End
   Begin VB.CommandButton cmdUpd 
      Caption         =   "&Apply"
      Height          =   480
      Left            =   9660
      TabIndex        =   10
      ToolTipText     =   "Update This Operation And Apply Changes"
      Top             =   960
      Width           =   875
   End
   Begin VB.TextBox txtScr 
      Height          =   315
      Left            =   4260
      TabIndex        =   7
      Tag             =   "1"
      ToolTipText     =   "Scrap Quantity (Rejected)"
      Top             =   4020
      Width           =   1095
   End
   Begin VB.TextBox txtRwk 
      Height          =   315
      Left            =   1140
      TabIndex        =   6
      Tag             =   "1"
      ToolTipText     =   "Quantity For Rework (Rejected)"
      Top             =   4020
      Width           =   1095
   End
   Begin VB.ComboBox txtIns 
      Height          =   315
      Left            =   1140
      Sorted          =   -1  'True
      TabIndex        =   8
      ToolTipText     =   "Enter Inspector Or Select From List"
      Top             =   4380
      Width           =   2295
   End
   Begin VB.ComboBox txtAcd 
      Height          =   315
      Left            =   3060
      TabIndex        =   3
      Tag             =   "4"
      Top             =   3660
      Width           =   1095
   End
   Begin VB.TextBox txtQty 
      Height          =   315
      Left            =   4260
      TabIndex        =   4
      Tag             =   "1"
      ToolTipText     =   "Accepted Quantity"
      Top             =   3660
      Width           =   1095
   End
   Begin VB.ComboBox cmbJmp 
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   660
      Sorted          =   -1  'True
      TabIndex        =   24
      TabStop         =   0   'False
      Tag             =   "8"
      ToolTipText     =   "Jump To Operation"
      Top             =   7380
      Width           =   2445
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Close"
      Height          =   480
      Left            =   9660
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin VB.TextBox txtNte 
      Height          =   285
      Left            =   5820
      TabIndex        =   11
      Tag             =   "2"
      Top             =   7380
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ComboBox cmbShp 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1140
      TabIndex        =   1
      Tag             =   "8"
      ToolTipText     =   "Shop"
      Top             =   3300
      Width           =   1815
   End
   Begin VB.ComboBox cmbWcn 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1140
      TabIndex        =   2
      Tag             =   "8"
      ToolTipText     =   "Work Center"
      Top             =   3660
      Width           =   1815
   End
   Begin VB.TextBox txtCmt 
      Enabled         =   0   'False
      Height          =   1695
      Left            =   1140
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      TabStop         =   0   'False
      Tag             =   "9"
      ToolTipText     =   "Comment (5120 Max)"
      Top             =   4740
      Width           =   9345
   End
   Begin VB.TextBox txtOpn 
      Enabled         =   0   'False
      Height          =   285
      Left            =   180
      TabIndex        =   12
      TabStop         =   0   'False
      Tag             =   "1"
      ToolTipText     =   "Operation "
      Top             =   3300
      Width           =   495
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   60
      Top             =   7005
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   8115
      FormDesignWidth =   10680
   End
   Begin MSFlexGridLib.MSFlexGrid Grd 
      Height          =   2955
      Left            =   1080
      TabIndex        =   0
      ToolTipText     =   "Click To Select Or Scroll And Press Enter"
      Top             =   0
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   5212
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      ScrollBars      =   2
      SelectionMode   =   1
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Comments"
      Height          =   255
      Index           =   10
      Left            =   180
      TabIndex        =   37
      Top             =   4800
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tool List"
      ForeColor       =   &H00400000&
      Height          =   255
      Index           =   12
      Left            =   180
      TabIndex        =   34
      Top             =   6540
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Org Quantity    "
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
      Left            =   5460
      TabIndex        =   29
      Top             =   3060
      Width           =   1095
   End
   Begin VB.Label lblOrig 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5460
      TabIndex        =   28
      ToolTipText     =   "Beginning Run Quantity"
      Top             =   3300
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Reduction"
      Height          =   285
      Index           =   9
      Left            =   3060
      TabIndex        =   27
      ToolTipText     =   "Amount Of Scrap"
      Top             =   4020
      Width           =   825
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Rework"
      Height          =   285
      Index           =   8
      Left            =   180
      TabIndex        =   26
      Top             =   4020
      Width           =   825
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Jump"
      Height          =   285
      Index           =   7
      Left            =   180
      TabIndex        =   25
      Top             =   7380
      Width           =   705
   End
   Begin VB.Label lblRun 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2940
      TabIndex        =   23
      Top             =   7740
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblPrt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   660
      TabIndex        =   22
      Top             =   7740
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Inspector"
      Height          =   255
      Index           =   5
      Left            =   180
      TabIndex        =   20
      Top             =   4380
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Notes"
      Height          =   255
      Index           =   4
      Left            =   3780
      TabIndex        =   19
      Top             =   7740
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Req/Acc Qty    "
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
      Left            =   4260
      TabIndex        =   18
      Top             =   3060
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sch/Act Date"
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
      Left            =   3060
      TabIndex        =   17
      Top             =   3060
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Shop/Work Center          "
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
      Index           =   1
      Left            =   1140
      TabIndex        =   16
      Top             =   3060
      Width           =   1815
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Op No"
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
      Index           =   0
      Left            =   180
      TabIndex        =   15
      Top             =   3060
      Width           =   495
   End
   Begin VB.Label lblReq 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   4260
      TabIndex        =   14
      ToolTipText     =   "Most Current Quantity Remaining"
      Top             =   3300
      Width           =   1095
   End
   Begin VB.Label lblSch 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3060
      TabIndex        =   13
      Top             =   3300
      Width           =   1035
   End
End
Attribute VB_Name = "ShopSHe03b"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'10/6/03 Added Rework and scrap
'7/8/04 Smoothed grid scrolling
'1/31/05 Added Trap for empty comment
'1/6/05 Removed Height for no comments (skews the dialog)
Option Explicit

'grid columns
Private Const OPCOLUMN_Number = 0
Private Const OPCOLUMN_Shop = 1
Private Const OPCOLUMN_WorkCenter = 2
Private Const OPCOLUMN_Comment = 3

Dim AdoOps As ADODB.Command
Dim AdoParameter1 As ADODB.Parameter
Dim ADOParameter2 As ADODB.Parameter
Dim AdoParameter3 As ADODB.Parameter


Dim bChanged As Byte
Dim bOnLoad As Byte
Dim bOpComplete As Byte
Dim iIndex As Integer
Dim iTotalOps As Integer
Dim iOpCur As Integer

Dim cOrigQty As Currency
Dim cRunqty As Currency

Dim sShop As String
Dim sCenter As String

Dim sOldShop As String
Dim sOldCenter As String
Dim sPartNumber As String
Dim sComments As String

Dim iOpno(300) As Integer

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub LastOp()
   Dim iRow As Integer
   iRow = Grd.row
   iRow = iRow - 1
   If iRow < 1 Then iRow = 1
   Grd.row = iRow
   Grd.Col = OPCOLUMN_Number
   txtOpn = Grd.Text
   GetThisOp
   
End Sub

Private Sub NextOp()
   Dim iRow As Integer
   iRow = Grd.row
   iRow = iRow + 1
   If iRow > Grd.Rows - 1 Then iRow = Grd.Rows - 1
   Grd.row = iRow
   Grd.Col = OPCOLUMN_Number
   txtOpn = Grd.Text
   GetThisOp
   
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtAcd = Format(ES_SYSDATE, "mm/dd/yy")
   
End Sub


Private Sub cmbJmp_Click()
   '    Dim b As Byte
   '    b = CheckChange()
   '    If b = 1 Then
   '        If cmbJmp.ListIndex >= 0 Then
   '            iIndex = cmbJmp.ListIndex + 1
   '            txtOpn = Format(iOpno(iIndex), "000")
   '            GetThisOp
   '        End If
   '    End If
   
End Sub


Private Sub cmbJmp_LostFocus()
   On Error Resume Next
   If Trim(cmbJmp) = "" Then cmbJmp = cmbJmp.List(iIndex - 1)
   
End Sub


Private Sub cmbShp_Change()
   bChanged = 1
   
End Sub

Private Sub cmbShp_Click()
   If cmbShp <> sOldShop Then FillWorkCenters
   
End Sub


Private Sub cmbShp_LostFocus()
   cmbShp = CheckLen(cmbShp, 12)
   FindShop
   sShop = Compress(cmbShp)
   If cmbShp <> sOldShop Then FillWorkCenters
   
End Sub

Private Sub cmbWcn_Change()
   bChanged = 1
   
End Sub

Private Sub cmbWcn_LostFocus()
   cmbWcn = CheckLen(cmbWcn, 12)
   sCenter = Compress(cmbWcn)
   
End Sub

Private Sub cmdCan_Click()
   Dim b As Byte
   b = CheckChange(1)
   If b = 1 Then
      SetMoCurrentOp
      Unload Me
   End If
   
End Sub

Private Sub cmdEnd_Click()
   Dim b As Byte
   b = CheckChange()
   If b = 1 Then
      txtOpn = Format(iOpno(iTotalOps), "000")
      GetThisOp
   End If
   
End Sub


Private Sub cmdFst_Click()
   Dim b As Byte
   b = CheckChange()
   If b = 1 Then
      txtOpn = Format(iOpno(1), "000")
      GetThisOp
   End If
   
End Sub

Private Sub cmdCompleteAll_Click()
    Dim sMsg As String
    
    sMsg = "Complete all operations with full MO quantity on today's date?"
    If MsgBox(sMsg, ES_NOQUESTION, Caption) <> vbYes Then
        Exit Sub
    End If
    
   'don't allow completion of all operations unless time charges are closed
   Dim rdoCharges As ADODB.Recordset
   sSql = "select distinct rtrim(PREMFSTNAME) + ' ' + rtrim(PREMLSTNAME) as Name," & vbCrLf _
    & "ISMOSTART, ISSHOP, ISWCNT, IsNull(ISSURUN,'R') as SURUN," & vbCrLf _
    & "(select cast(ISOP as VARCHAR(4)) + ' '  from IstcTable chg2" & vbCrLf _
    & " where chg2.ISMO = chg1.ISMO and chg2.ISRUN = chg1.ISRUN" & vbCrLf _
    & " order by ISOP" & vbCrLf _
    & " for XML PATH('')" & vbCrLf _
    & ") as Ops" & vbCrLf _
    & "from IstcTable chg1" & vbCrLf _
    & "join EmplTable emp on emp.PREMNUMBER = chg1.ISEMPLOYEE" & vbCrLf _
    & "where ISMO = '" & lblPrt.Caption & "' and ISRUN = " & lblRun.Caption

    bSqlRows = clsADOCon.GetDataSet(sSql, rdoCharges)
    'clsadocon.GetQuerySet
    If bSqlRows Then
        sMsg = "The following time charges must be closed in order to proceed:" & vbCrLf
        With rdoCharges
            While Not .EOF
                sMsg = sMsg & !Name & " " & !Ops & vbCrLf
                .MoveNext
            Wend
        End With
        sMsg = sMsg & "Close time charges as of current time and proceed?"
        If MsgBox(sMsg, vbYesNo) <> vbYes Then
            Exit Sub
        End If
        
        Dim cmd As New ADODB.Command
        Dim rs As ADODB.Recordset
        cmd.CommandText = "exec CompleteAllOps ?, ?, ? OUT"
        
        cmd.Parameters.Append cmd.CreateParameter("PartRef", adVarChar, adParamInput, 30)
        cmd.Parameters("PartRef").Value = lblPrt.Caption
        
        cmd.Parameters.Append cmd.CreateParameter("RunNo", adInteger, adParamInput)
        cmd.Parameters("RunNo") = lblRun.Caption
        
        cmd.Parameters.Append cmd.CreateParameter("NeedJournal", adInteger, adParamOutput)
        
        'bSqlRows = clsADOCon.GetQuerySet(rs, cmd)
        
        clsADOCon.GetQuerySet rs, cmd, ES_KEYSET, True
        'cmd.Execute
        If cmd.Parameters("NeedJournal") <> 0 Then
            MsgBox "There is no open time journal for the required time charge dates.  Unable to proceed", vbOKOnly
            Exit Sub
        End If
    End If
    
    'First complete time charges
'    rdoCharges.MoveFirst
'    With rdoCharges
'        While Not .EOF
'
'        Loop
'    End With
    
      If CompleteAll Then
         MsgBox "All operations have been completed", vbInformation, Caption
         Unload Me
      Else
         MsgBox "Completion of all operations failed.  Please perform operation completions individually.", vbExclamation, Caption
         FillOps 'some may have completed
      End If
      
   
   
   
   
End Sub

Private Function CompleteAll() As Boolean
   'complete all operations with today's date for full MO quantity
   'return true if successful
   
   MouseCursor ccHourglass
   
   Dim row As Integer, completionDate As String, qty As Currency
   Dim PartRef As String, Runno As Integer, opNo As Integer
   completionDate = Format(ES_SYSDATE, "mm/dd/yy")
   PartRef = lblPrt.Caption
   Runno = CInt(lblRun)
   qty = CInt(lblOrig.Caption)
   On Error GoTo whoops
   clsADOCon.BeginTrans
   clsADOCon.ADOErrNum = 0
   
'   For row = 1 To Grd.Rows - 1
'      'Debug.Print CInt(Grd.TextMatrix(row, OPCOLUMN_Number))
'      opNo = CInt(Grd.TextMatrix(row, OPCOLUMN_Number))
'      sSql = "UPDATE RnopTable SET " _
'             & "OPCOMPDATE='" & completionDate & "', " _
'             & "OPYIELD=" & qty & ", " _
'             & "OPCOMPLETE=1" & vbCrLf _
'             & "WHERE OPREF='" & PartRef & "' AND " _
'             & "OPRUN=" & Runno & " AND OPNO=" & opNo
'      clsADOCon.ExecuteSql sSql
'   Next

    'complete incomplete operations all at once.  also set value of OPACCEPT
    sSql = "UPDATE RnopTable SET " _
        & "OPCOMPDATE='" & completionDate & "', " _
        & "OPYIELD=" & qty & ", " _
        & "OPACCEPT=" & qty & ", " _
        & "OPCOMPLETE=1" & vbCrLf _
        & "WHERE OPREF='" & PartRef & "' AND " _
        & "OPRUN=" & Runno & " AND OPCOMPLETE=0"
    clsADOCon.ExecuteSql sSql
    
   clsADOCon.CommitTrans
   MouseCursor ccDefault
   CompleteAll = True
   Exit Function
   
whoops:
   clsADOCon.RollbackTrans
   MouseCursor ccDefault
   CompleteAll = False
   
End Function

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 4103
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub cmdLst_Click()
   Dim A As Integer
   Dim b As Byte
   b = CheckChange()
   If b = 1 Then
      iIndex = iIndex - 1
      If iIndex < 1 Then iIndex = 1
      txtOpn = Format(iOpno(iIndex), "000")
      GetThisOp
   End If
   If Grd.row < 6 Then
      Grd.TopRow = 1
   Else
      A = Grd.row Mod 4
      If A = 2 Then Grd.TopRow = Grd.row - 4
   End If
   
End Sub

Private Sub cmdNxt_Click()
   Dim A As Integer
   Dim b As Byte
   b = CheckChange()
   If b = 1 Then
      iIndex = iIndex + 1
      If iIndex > iTotalOps Then iIndex = iTotalOps
      txtOpn = Format(iOpno(iIndex), "000")
      GetThisOp
   End If
   Grd.Col = OPCOLUMN_Number
   If Grd.row < 6 Then
      Grd.TopRow = 1
   Else
      A = Grd.row Mod 4
      If A = 2 Then Grd.TopRow = Grd.row - 1
   End If
   
End Sub

Private Sub cmdUpd_Click()
   Dim b As Byte
   b = CheckOpQuantity()
   If b = 0 Then
      If MsgBox("The Rejected And Accepted Quantities Are" & vbCr _
         & "Greater Than The Available Quantity." & vbCrLf _
         & "Do you wish to proceed anyway?", _
         vbYesNo + vbQuestion, Caption) = vbNo Then Exit Sub
   End If
   If optCom.Value = vbUnchecked Then txtAcd = ""
   UpdateOp
   
End Sub

Private Sub Form_Activate()
   If bOnLoad Then
      GetMoreRunInfo
      FillInspectors
      FillOps
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   'FormLoad Me, ES_DONTLIST
   'Move ShopSHe03a.Left + 500, ShopSHe03a.Top + 600
   bOnLoad = 1
   FormatControls
   If ShopSHe03a.OptCmt = vbChecked Then
      txtCmt.Visible = True
   Else
      txtCmt.Visible = False
      lblCurList.Top = 3250
      lblLst.Top = 3600
      z1(12).Top = 3250
      cmdNxt.Top = 3250
      cmdLst.Top = 3250
   End If
   sSql = "SELECT OPREF,OPRUN,OPNO,OPSHOP,OPCENTER," _
          & "OPSCHEDDATE,OPNOTES,OPCOMPDATE,OPINSP,OPYIELD," _
          & "OPCOMPLETE,OPACCEPT,OPREWORK,OPSCRAP,OPTOOLLIST,OPCOMT " _
          & "FROM RnopTable WHERE OPREF= ? AND OPRUN= ? AND OPNO= ?"

   Set AdoOps = New ADODB.Command
   AdoOps.CommandText = sSql
   
   Set AdoParameter1 = New ADODB.Parameter
   AdoParameter1.Type = adChar
   AdoParameter1.SIZE = 30
   
   Set ADOParameter2 = New ADODB.Parameter
   ADOParameter2.Type = adInteger
   
   Set AdoParameter3 = New ADODB.Parameter
   AdoParameter3.Type = adSmallInt
   
   AdoOps.Parameters.Append AdoParameter1
   AdoOps.Parameters.Append ADOParameter2
   AdoOps.Parameters.Append AdoParameter3

   With Grd
      .ColAlignment(0) = 0
      .ColAlignment(1) = 0
      .ColAlignment(2) = 0
      .row = 0
      
      .Col = OPCOLUMN_Number
      .Text = "Op No"
      .ColWidth(0) = 750
      
      .Col = OPCOLUMN_Shop
      .Text = "Shop"
      .ColWidth(1) = 800
      
      .Col = OPCOLUMN_WorkCenter
      .Text = "Work Ctr"
      .ColWidth(2) = 850
      
      .Col = OPCOLUMN_Comment
      .Text = "Comment"
      .ColWidth(3) = 5800
      
      .Col = OPCOLUMN_Number
   End With
   sPartNumber = ShopSHe03a.cmbPrt
   sPartNumber = Compress(sPartNumber)
   lblPrt = sPartNumber
   lblRun = ShopSHe03a.cmbRun
   'bOnLoad = 1
   
End Sub


Private Sub UpdateOp()
   Dim cYield As Currency
   Dim sCompDate As String
   MouseCursor 13
   If Len(Trim(cmbShp)) = "" Then
      MsgBox "Requires A Valid Shop.", vbExclamation, Caption
      Exit Sub
   End If
   If optCom Then
      cYield = Val(txtQty)
      sCompDate = "'" & txtAcd & "'"
   Else
      cYield = 0
      sCompDate = "Null"
   End If
   
   sShop = Compress(cmbShp)
   sCenter = Compress(cmbWcn)
   On Error Resume Next
   clsADOCon.ADOErrNum = 0
   
   sSql = "UPDATE RnopTable SET " _
          & "OPSHOP='" & sShop & "'," _
          & "OPCENTER='" & sCenter & "'," _
          & "OPCOMPDATE=" & sCompDate & "," _
          & "OPNOTES='" & txtNte & "'," _
          & "OPINSP='" & txtIns & "'," _
          & "OPYIELD=" & cYield & "," _
          & "OPACCEPT=" & Val(cYield) & "," _
          & "OPREWORK=" & Val(txtRwk) & "," _
          & "OPSCRAP=" & Val(txtScr) & "," _
          & "OPCOMT='" & txtCmt & "'," _
          & "OPCOMPLETE=" & str(optCom.Value) & " " _
          & "WHERE OPREF='" & lblPrt & "' AND " _
          & "OPRUN=" & Val(lblRun) & " AND OPNO=" & Val(txtOpn)
   clsADOCon.ExecuteSql sSql
   bChanged = 0
   If clsADOCon.ADOErrNum = 0 Then UpdateMo
   MouseCursor 0
   FillOps
   Exit Sub
   
DiaErr1:
   sProcName = "updateop"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   ShopSHe03a.optLoaded = vbUnchecked
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   Set AdoParameter1 = Nothing
   Set ADOParameter2 = Nothing
   Set AdoParameter3 = Nothing
   Set AdoOps = Nothing
   Set ShopSHe03b = Nothing
   
End Sub


Private Sub GetThisOp()
   Dim RdoRps As ADODB.Recordset
   On Error GoTo DiaErr1
   AdoOps.Parameters(0).Value = lblPrt
   AdoOps.Parameters(1).Value = Val(lblRun)
   AdoOps.Parameters(2).Value = Val(txtOpn)
   bSqlRows = clsADOCon.GetQuerySet(RdoRps, AdoOps, ES_KEYSET, False, 1)
   If bSqlRows Then
      With RdoRps
         cmbShp = "" & Trim(!OPSHOP)
         cmbWcn = "" & Trim(!OPCENTER)
         sOldCenter = cmbWcn
         cmbWcn = GetCenter(cmbWcn)
         'txtCmt = "" & Trim(Left(!OPCOMT, 20))
         txtCmt = "" & Trim(!OPCOMT)
         lblOrig = Format(cOrigQty, ES_QuantityDataFormat)
         lblSch = "" & Format(!OPSCHEDDATE, "mm/dd/yy")
         txtRwk = Format(!OPREWORK, ES_QuantityDataFormat)
         txtScr = Format(!OPSCRAP, ES_QuantityDataFormat)
         txtNte = "" & Trim(!OPNOTES)
         txtIns = "" & Trim(!OPINSP)
         lblCurList = "" & Trim(!OPTOOLLIST)
         If lblCurList <> "" Then lblCurList = FindToolList(lblCurList, lblLst) _
            Else lblLst = ""
         If !OPCOMPLETE = 0 Then
            lblReq = Format(cRunqty, ES_QuantityDataFormat)
            cmbShp.Enabled = True
            cmbWcn.Enabled = True
            cmbShp.ForeColor = ES_BLUE
            cmbWcn.ForeColor = ES_BLUE
            txtRwk.Enabled = True
            txtScr.Enabled = True
            txtAcd.Enabled = True
            txtQty.Enabled = True
            txtNte.Enabled = True
            txtCmt.Enabled = True
            optCom.Value = vbUnchecked
            txtQty = Format(cRunqty, ES_QuantityDataFormat)
         Else
            lblReq = Format(!OPACCEPT, ES_QuantityDataFormat)
            cmbShp.Enabled = False
            cmbWcn.Enabled = False
            cmbShp.ForeColor = vbGrayText
            cmbWcn.ForeColor = vbGrayText
            txtRwk.Enabled = False
            txtScr.Enabled = False
            txtAcd.Enabled = False
            txtQty.Enabled = False
            txtNte.Enabled = False
            txtCmt.Enabled = False
            optCom.Value = vbChecked
            txtAcd = "" & Format(!OPCOMPDATE, "mm/dd/yy")
            txtQty = Format(0 + !OPYIELD, ES_QuantityDataFormat)
         End If
         bOpComplete = !OPCOMPLETE
         ClearResultSet RdoRps
      End With
   End If
   On Error Resume Next
   Set RdoRps = Nothing
   'grd.row = iIndex
   'grd.Col = OPCOLUMN_Number
   'grd.SetFocus
   FindShop
   sOldShop = cmbShp
   'FillWorkCenters
   bChanged = 0
   Exit Sub
   
DiaErr1:
   sProcName = "getthisop"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub FillOps()
   Dim RdoFil As ADODB.Recordset
   Dim iRows As Integer
   Dim sShop As String
   Dim sCenter As String
   Dim sComment As String
   Erase iOpno
   Grd.Rows = 2
   iTotalOps = 0
   iIndex = 1
   On Error GoTo DiaErr1
   If ShopSHe03a.optOps.Value = vbChecked Then
      sSql = "SELECT OPREF,OPNO,OPRUN,OPSHOP,OPCENTER,OPCOMPLETE,OPCOMT FROM RnopTable WHERE " _
             & "(OPREF='" & lblPrt & "' AND OPRUN=" & lblRun & " AND OPCOMPLETE=0)"
   Else
      sSql = "SELECT OPREF,OPNO,OPRUN,OPSHOP,OPCENTER,OPCOMPLETE,OPCOMT FROM RnopTable WHERE " _
             & "(OPREF='" & lblPrt & "' AND OPRUN=" & lblRun & ")"
   End If
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoFil, ES_KEYSET)
   If bSqlRows Then
      With RdoFil
         txtOpn = Format(!opNo, "000")
         Do Until .EOF
            iTotalOps = iTotalOps + 1
            iOpno(iTotalOps) = !opNo
            cmbJmp.AddItem Format(!opNo, "000") & " "
            iRows = iRows + 1
            If iRows > 1 Then Grd.Rows = Grd.Rows + 1
            Grd.row = iRows
            
            'grd.col = OPCOLUMN_Number
            'grd.Text = Format(!opNo, "000")
            Grd.TextMatrix(iRows, OPCOLUMN_Number) = Format(!opNo, "000")
            
            'grd.col = OPCOLUMN_Shop
            'sShop = GetRoutShop("" & Trim(!OPSHOP))
            'grd.Text = sShop
            Grd.TextMatrix(iRows, OPCOLUMN_Shop) = GetRoutShop("" & Trim(!OPSHOP))
            
            'grd.col = OPCOLUMN_WorkCenter
            'sCenter = GetRoutCenter("" & Trim(!OPCENTER))
            'grd.Text = sCenter
            Grd.TextMatrix(iRows, OPCOLUMN_WorkCenter) = GetRoutCenter("" & Trim(!OPCENTER))
            
            'grd.col = OPCOLUMN_Comment
            'sComment = "" & Trim(Left(!OPCOMT, 20))
            'grd.Text = sComment
            Grd.TextMatrix(iRows, OPCOLUMN_Comment) = "" & Replace(Trim(Left(!OPCOMT, 55)), vbCrLf, " ")
            
            .MoveNext
         Loop
         DoEvents
         ClearResultSet RdoFil
      End With
      If cmbJmp.ListCount > 0 Then cmbJmp = cmbJmp.List(0)
      Grd.row = 1
      Grd.Col = OPCOLUMN_Number
   Else
      MsgBox "There Are No Incomplete Operations For This MO.", _
         vbInformation, Caption
      Unload Me
      Exit Sub
   End If
   
   Set RdoFil = Nothing
   
   sSql = "Qry_FillShops "
   LoadComboBox cmbShp
   GetThisOp
   Exit Sub
   
DiaErr1:
   sProcName = "fillops"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Private Sub FillWorkCenters()
   cmbWcn.Clear
   If cmbShp = "" Then cmbShp = sOldShop
   On Error GoTo DiaErr1
   sSql = "Qry_FillWorkCenters '" & Compress(cmbShp) & "'"
   LoadComboBox cmbWcn
   If cmbWcn.ListCount > 0 Then cmbWcn = cmbWcn.List(0)
   sOldShop = cmbShp
   cmbWcn = GetCenter(cmbWcn)
   Exit Sub
   
DiaErr1:
   sProcName = "fillworkcenters"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Private Sub GetMoreRunInfo()
   Dim RdoInf As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT RUNREF,RUNNO,RUNQTY,RUNOPCUR,RUNREMAININGQTY " _
          & "FROM RunsTable WHERE " _
          & "RUNREF='" & lblPrt & "' AND RUNNO=" & Val(lblRun) & " "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoInf)
   If bSqlRows Then
      On Error Resume Next
      With RdoInf
         cOrigQty = !RUNQTY
         cRunqty = !RUNREMAININGQTY
         iOpCur = !RUNOPCUR
      End With
   End If
   '3/4/05 Nulls
   sSql = "UPDATE RnopTable SET OPCOMT='' WHERE OPCOMT IS NULL"
   clsADOCon.ExecuteSql sSql
   Set RdoInf = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getmorerun"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub



Private Sub Grd_Click()
   'grd.col = OPCOLUMN_Number
   iIndex = Grd.row
   'txtOpn = grd.Text
   txtOpn = Grd.TextMatrix(Grd.row, OPCOLUMN_Number)
   GetThisOp
End Sub

Private Sub Grd_EnterCell()
   If bOnLoad = 0 Then
      Grd.Col = OPCOLUMN_Number
      iIndex = Grd.row
      txtOpn = Grd.Text
      GetThisOp
   End If
End Sub

Private Sub Grd_GotFocus()
   Grd.Col = OPCOLUMN_Number
   
End Sub

Private Sub grd_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Or KeyAscii = vbKeySpace Then
      iIndex = Grd.row
      Grd.Col = OPCOLUMN_Number
      txtOpn = Grd.Text
      GetThisOp
   End If
   
End Sub


Private Sub optCom_Click()
   If bOnLoad = 0 Then
      CheckOps
      bChanged = 1
   Else
      bOnLoad = 0
   End If
   
End Sub


Private Function TrimComment(sComment As String)
   Dim n As String
   On Error GoTo DiaErr1
   If Len(sComment) > 0 Then
      sComment = Replace(sComment, vbCrLf, " ")
      sComment = Replace(sComment, vbCr, " ")
      sComment = Replace(sComment, vbLf, " ")
      If Len(sComment) > 20 Then sComment = Left(sComment, 20)
   End If
   TrimComment = sComment
   Exit Function
   
DiaErr1:
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   On Error GoTo 0
   
End Function

Private Sub optCom_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then LastOp
   If KeyCode = vbKeyPageDown Then NextOp
   
End Sub


Private Sub txtAcd_Change()
   bChanged = 1
   
End Sub

Private Sub txtAcd_DropDown()
   ShowCalendar Me
   
End Sub

Private Sub txtAcd_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then LastOp
   If KeyCode = vbKeyPageDown Then NextOp
   
End Sub

Private Sub txtAcd_LostFocus()
   If Len(Trim(txtAcd)) Then txtAcd = CheckDate(txtAcd)
   
End Sub

Private Sub txtCmt_Change()
   bChanged = 1
   
End Sub

Private Sub txtCmt_LostFocus()
   txtCmt = CheckLen(txtCmt, 5120)
   txtCmt = StrCase(txtCmt, ES_FIRSTWORD)
   txtCmt = ReplaceString(txtCmt)
'   Grd.Col = OPCOLUMN_Comment
'   Grd.Text = Replace(Left$(txtCmt, 55), vbCrLf, "")
   Grd.TextMatrix(Grd.row, OPCOLUMN_Comment) = Replace(Left$(txtCmt, 55), vbCrLf, "")
   
End Sub


Private Sub txtIns_Change()
   bChanged = 1
   
End Sub

Private Sub txtIns_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then LastOp
   If KeyCode = vbKeyPageDown Then NextOp
   
End Sub

Private Sub txtIns_LostFocus()
   txtIns = CheckLen(txtIns, 30)
   StrCase (txtIns)
   
End Sub

Private Sub txtNte_Change()
   bChanged = 1
   
End Sub

Private Sub txtNte_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then LastOp
   If KeyCode = vbKeyPageDown Then NextOp
   
End Sub

Private Sub txtNte_LostFocus()
   txtNte = CheckLen(txtNte, 6)
   
End Sub


Private Sub txtQty_Change()
   bChanged = 1
   
End Sub

Private Sub txtQty_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then LastOp
   If KeyCode = vbKeyPageDown Then NextOp
   
End Sub

Private Sub txtQty_LostFocus()
   txtQty = CheckLen(txtQty, 9)
   txtQty = Format(Abs(Val(txtQty)), ES_QuantityDataFormat)
   
End Sub



Private Sub CheckOps()
   If optCom.Value = vbChecked Then
      cmbShp.Enabled = False
      cmbWcn.Enabled = False
      cmbShp.ForeColor = vbGrayText
      cmbWcn.ForeColor = vbGrayText
      txtRwk.Enabled = False
      txtScr.Enabled = False
      txtAcd.Enabled = False
      txtQty.Enabled = False
   Else
      cmbShp.Enabled = True
      cmbWcn.Enabled = True
      cmbShp.ForeColor = ES_BLUE
      cmbWcn.ForeColor = ES_BLUE
      txtRwk.Enabled = True
      txtScr.Enabled = True
      txtAcd.Enabled = True
      If Trim(txtAcd) = "" Then txtAcd = Format(ES_SYSDATE, "mm/dd/yy")
      txtQty.Enabled = True
   End If
   
End Sub

Private Function GetCenter(sNewCenter As String) As String
   Dim RdoCnt As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "Qry_GetWorkCenter '" & Compress(sNewCenter) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCnt, ES_FORWARD)
   If bSqlRows Then
      With RdoCnt
         GetCenter = Trim(!WCNNUM)
         ClearResultSet RdoCnt
      End With
   Else
      GetCenter = ""
   End If
   Set RdoCnt = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getcenter"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub FillInspectors()
   Dim RdoIns As ADODB.Recordset
   Dim sIns As String
   sSql = "SELECT INSID,INSFIRST,INSMIDD,INSLAST FROM RinsTable" & vbCrLf _
      & "WHERE INSACTIVE = 1" & vbCrLf _
      & "ORDER BY INSFIRST, INSMIDD, INSLAST"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoIns, ES_FORWARD)
   If bSqlRows Then
      With RdoIns
         Do Until .EOF
            sIns = "" & Trim(!INSFIRST)
            If Len(Trim(!INSMIDD)) Then sIns = sIns & " " & Trim(!INSMIDD) & "."
            sIns = sIns & Trim(!INSLAST)
            AddComboStr txtIns.hwnd, sIns
            .MoveNext
         Loop
         ClearResultSet RdoIns
      End With
   End If
   Set RdoIns = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "fillinsp"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

'Set the current op

Private Sub SetMoCurrentOp()
   On Error GoTo DiaErr1
   Dim RdoLst As ADODB.Recordset
'   sSql = "SELECT OPREF,OPRUN,OPNO,OPCOMPLETE FROM RnopTable " _
'          & "WHERE (OPREF='" & lblPrt & "' AND OPRUN=" & Val(lblRun) & " " _
'          & "AND OPCOMPLETE=0) ORDER BY OPNO "
          
   sSql = "SELECT OPREF,OPRUN,OPNO,OPCOMPLETE, OPCOMPDATE,OPNOCOMP FROM RnopTable," _
          & "   (SELECT TOP(1) OPNO as OPNOCOMP FROM RnopTable " _
          & "            WHERE OPREF='" & lblPrt & "' AND OPRUN=" & Val(lblRun) & " AND OPCOMPLETE = 1" _
          & "               ORDER BY OPNO DESC) as f " _
          & "   WHERE (OPREF='" & lblPrt & "' AND OPRUN=" & Val(lblRun) & " AND OPCOMPLETE=0 AND OPNO > OPNOCOMP)"
          
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoLst, ES_FORWARD)
   If bSqlRows Then
      With RdoLst
         iOpCur = !opNo
         ClearResultSet RdoLst
      End With
   Else
      sSql = "SELECT MAX(OPNO) FROM RnopTable " _
             & "WHERE (OPREF='" & lblPrt & "' AND OPRUN=" & Val(lblRun) & ")"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoLst, ES_FORWARD)
      If bSqlRows Then
         With RdoLst
            If Not IsNull(.Fields(0)) Then _
                          iOpCur = .Fields(0)
            ClearResultSet RdoLst
         End With
      End If
   End If
   
   Set RdoLst = Nothing

   sSql = "UPDATE runsTable set RUNOPCUR = MinOPno FROM " _
         & "(select MIN(OPNO) as MinOpno,  opref, Oprun FROM rnopTable " _
            & "where OPCOMPLETE = 0 AND opref = '" & lblPrt & "' and " _
               & " OPRUN=" & lblRun & " GROUP BY opref, Oprun) as f  " _
            & " WHERE f.OPREF = runref and f.OPRUN = runno " _
            & " And runref = '" & lblPrt & "' AND runno = " & lblRun _
            & " AND MinOPno <> RUNOPCUR"
   
   clsADOCon.ExecuteSql sSql
 

'   sSql = "UPDATE RunsTable SET RUNOPCUR=" & Trim(str(iOpCur)) & " " _
'          & "WHERE RUNREF='" & lblPrt & "' AND RUNNO=" & lblRun & " "
          
'   clsADOCon.ExecuteSQL sSql
   Exit Sub
   
DiaErr1:
   sProcName = "SetMoCurrent"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub txtRwk_Change()
   bChanged = 1
   
End Sub

Private Sub txtRwk_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then LastOp
   If KeyCode = vbKeyPageDown Then NextOp
   
End Sub


Private Sub txtRwk_LostFocus()
   txtRwk = CheckLen(txtRwk, 9)
   txtRwk = Format(Abs(Val(txtRwk)), ES_QuantityDataFormat)
   
End Sub


Private Sub txtScr_Change()
   bChanged = 1
   
End Sub

Private Sub txtScr_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then LastOp
   If KeyCode = vbKeyPageDown Then NextOp
   
End Sub



'10/6/03

Private Function CheckOpQuantity() As Byte
   Dim sAvail As Currency
   Dim sAccept As Currency
   Dim sReject As Currency
   
   sAvail = Val(lblReq)
   sAccept = Val(txtQty)
   sReject = Val(txtRwk) + Val(txtScr) + sAccept
   If sReject > sAvail Then CheckOpQuantity = 0 Else _
                                              CheckOpQuantity = 1
   
End Function

Private Sub txtScr_LostFocus()
   txtScr = CheckLen(txtScr, 9)
   txtScr = Format(Abs(Val(txtScr)), ES_QuantityDataFormat)
   
End Sub



'10/6/03

Private Sub UpdateMo()
   Dim RdoQty As ADODB.Recordset
   Dim cRework As Currency
   Dim cReject As Currency
   Dim cScrap As Currency
   
   sSql = "SELECT SUM(OPREWORK),SUM(OPSCRAP) FROM RnopTable " _
          & "WHERE OPREF='" & lblPrt & "' AND OPRUN=" & Val(lblRun) & " "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoQty, ES_FORWARD)
   If bSqlRows Then
      With RdoQty
         If Not IsNull(.Fields(0)) Then
            cRework = .Fields(0)
         End If
         If Not IsNull(.Fields(1)) Then
            cScrap = .Fields(1)
         End If
         ClearResultSet RdoQty
      End With
   End If
   cReject = cRework + cScrap
   sSql = "UPDATE RunsTable SET RUNREMAININGQTY=RUNQTY-" _
          & cReject & ",RUNSCRAP=" & cScrap & ",RUNREWORK=" & cRework & " " _
          & "WHERE RUNREF='" & lblPrt & "' AND " _
          & "RUNNO=" & Val(lblRun) & " "
   clsADOCon.ExecuteSql sSql
   SysMsg "Operation Updated.", True
   Set RdoQty = Nothing
   
End Sub

Private Function CheckChange(Optional ExitForm As Byte) As Byte
   Dim bResponse As Byte
   If bChanged Then
      If ExitForm = 0 Then
         bResponse = MsgBox("The Data Has Changed. Change Operation" & vbCr _
                     & "Without Updating?", ES_NOQUESTION, Caption)
      Else
         bResponse = MsgBox("The Data Has Changed. Exit " _
                     & "Without Updating?", ES_NOQUESTION, Caption)
      End If
      If bResponse = vbYes Then CheckChange = 1 Else CheckChange = 0
   Else
      CheckChange = 1
   End If
   
End Function

