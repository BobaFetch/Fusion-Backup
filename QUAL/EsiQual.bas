Attribute VB_Name = "EsiQual"
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'Reject Tags changed to Inspection Reports 9/23/02
'Characteristic Codes changed to Discrepancy Codes 9/23/02
'4/17/03 First Article Inspection
'3/29/05 removed JetDb references
'8/8/05 Corrected KeySet clearing
'10/31/05 Added Cur.CurrentGroup to OpenFavorite. Opens appropriate tab
'         when called from Recent/Favorites and closed.
'1/12/06  Completed renaming dialogs to be consistent with Fina
'4/19/06  See UpdateTables for added list of stored procedures
'4/26/06  Changed BackColor Property of SPC Functions
'5/31/06 buildKeys/Convert Rej and Fai tables
'6/23/06 Removed Threed32.OCX
'8/9/06 Removed SSTab32.OCX
'1/17/07 Added GetThisVendor and GetThisCustomer for reports
Option Explicit

Public Y As Byte
Public bFoundPart As Byte
Public sCurrForm As String
Public sLastType As String
Public sSelected As String

Public sFavorites(13) As String
Public sRecent(10) As String
Public sSession(50) As String
Public sPassedPart As String

'Column updates
Private RdoCol As ADODB.Recordset
Private AdoError As ADODB.Error

'Private ER As rdoError

'Type tUser
'    Adduser As Integer
'    Level   As Integer
'    Group1  As Integer
'    Group2  As Integer
'    Group3  As Integer
'    Group4  As Integer
'    Group5  As Integer
'    Group6  As Integer
'    Group7  As Integer
'    Group8  As Integer
'End Type
'Public User As tUser
'


'1/17/07 Added for reports

Public Sub GetThisVendor()
   Dim RdoRpt As ADODB.Recordset
   On Error GoTo modErr1
   sSql = "SELECT VEBNAME FROM VndrTable WHERE VEREF='" _
          & Compress(MdiSect.ActiveForm.cmbVnd) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRpt, ES_FORWARD)
   If bSqlRows Then
      MdiSect.ActiveForm.lblVEName = "" & Trim(RdoRpt!VEBNAME)
      ClearResultSet RdoRpt
   Else
      MdiSect.ActiveForm.lblVEName = "*** A Range Of Vendors Selected ***"
   End If
   Set RdoRpt = Nothing
   Exit Sub
modErr1:
   
   On Error GoTo 0
End Sub

'1/17/07 Added for reports

Public Function GetReportID(strRptName As String) As Integer
   Dim RdoRpt As ADODB.Recordset
   On Error GoTo DiaErr1
   
   If (strRptName = "") Then
      MsgBox ("Please Select Certification Report Type")
      GetReportID = 0
      Exit Function
   End If
   
   sSql = "SELECT ISNULL(REPORTID, 0) REPORTID FROM CertReports WHERE REPORTNAME = '" & strRptName & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRpt, ES_FORWARD)
   If bSqlRows Then
      
      With RdoRpt
         GetReportID = Trim(!REPORTID)
         ClearResultSet RdoRpt
      End With
   Else
      GetReportID = 0
   End If
   Set RdoRpt = Nothing
   MouseCursor 0
   Exit Function
   
DiaErr1:
   sProcName = "FillReports"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description

End Function

Public Sub GetThisCustomer(Optional ControlIsTextBox As Byte)
   Dim RdoRpt As ADODB.Recordset
   On Error GoTo modErr1
   If ControlIsTextBox = 1 Then
      sSql = "SELECT CUNAME FROM CustTable WHERE CUREF='" _
             & Compress(MdiSect.ActiveForm.txtCst) & "'"
   Else
      sSql = "SELECT CUNAME FROM CustTable WHERE CUREF='" _
             & Compress(MdiSect.ActiveForm.cmbCst) & "'"
   End If
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRpt, ES_FORWARD)
   If bSqlRows Then
      MdiSect.ActiveForm.lblCUName = "" & Trim(RdoRpt!CUNAME)
      ClearResultSet RdoRpt
   Else
      MdiSect.ActiveForm.lblCUName = "*** A Range Of Customers Selected ***"
   End If
   Set RdoRpt = Nothing
   Exit Sub
   
modErr1:
   On Error GoTo 0
   
End Sub

Private Function CheckConvErrors() As Byte
   Dim iColCounter As Integer
   CheckConvErrors = 0
   For Each AdoError In RdoCol.ActiveConnection.Errors
      If Left(AdoError.Description, 5) = "22003" Then
         iColCounter = iColCounter + 1
         CheckConvErrors = 1
      End If
   Next AdoError
   
End Function

'5/30/06 Rejection Tags (Keys too)

Private Sub ConvertRejTagTables()
   Dim bBadCol As Byte
   Dim sconstraint As String
   On Error Resume Next
   Err.Clear
   'Start RjhdTable
   'REJREC
   sSql = "sp_columns @table_name=RjhdTable,@column_name=REJREC"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
 
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then 'See Else
               sSql = "ALTER TABLE RjhdTable ALTER COLUMN REJREC dec(12,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RjhdTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE RjhdTable ALTER COLUMN REJREC DEC(12,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RjhdTable.REJREC'"
               clsADOCon.ExecuteSQL sSql
            Else
               GoTo Keys
            End If
         End If
      End If
   End With
   'REJREJ
   sSql = "sp_columns @table_name=RjhdTable,@column_name=REJREJ"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RjhdTable ALTER COLUMN REJREJ dec(12,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RjhdTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE RjhdTable ALTER COLUMN REJREJ DEC(12,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RjhdTable.REJREJ'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'REJACCT
   sSql = "sp_columns @table_name=RjhdTable,@column_name=REJACCT"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RjhdTable ALTER COLUMN REJACCT dec(12,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RjhdTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE RjhdTable ALTER COLUMN REJACCT DEC(12,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RjhdTable.REJACCT'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   
   'Begin RjitTable
   'RITQTY
   sSql = "sp_columns @table_name=RjitTable,@column_name=RITQTY"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RjitTable ALTER COLUMN RITQTY dec(12,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RjitTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE RjitTable ALTER COLUMN RITQTY DEC(12,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RjitTable.RITQTY'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'RITRWK
   sSql = "sp_columns @table_name=RjitTable,@column_name=RITRWK"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RjitTable ALTER COLUMN RITRWK dec(12,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RjitTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE RjitTable ALTER COLUMN RITRWK DEC(12,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RjitTable.RITRWK'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'RITSCRP
   sSql = "sp_columns @table_name=RjitTable,@column_name=RITSCRP"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RjitTable ALTER COLUMN RITSCRP dec(12,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RjitTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE RjitTable ALTER COLUMN RITSCRP DEC(12,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RjitTable.RITSCRP'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   
Keys:
   Err.Clear
   clsADOCon.ADOErrNum = 0
   
   sSql = "DROP INDEX RjhdTable.RejRef"
   clsADOCon.ExecuteSQL sSql
   If clsADOCon.ADOErrNum = 0 Then
      sSql = "ALTER TABLE RjhdTable ALTER COLUMN REJREF CHAR(12) NOT NULL"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "ALTER TABLE RjhdTable ADD Constraint PK_RjhdTable_REJREF PRIMARY KEY CLUSTERED (REJREF) " _
             & "WITH FILLFACTOR=80 "
      clsADOCon.ExecuteSQL sSql
      
      sSql = "ALTER TABLE RjitTable ADD CONSTRAINT FK_RjitTable_RjhdTable FOREIGN KEY (RITREF) References RjhdTable ON DELETE CASCADE ON UPDATE CASCADE "
      clsADOCon.ExecuteSQL sSql
   End If
   
End Sub


Private Function GetConstraint(sDescription As String) As String
   Dim bByte As Byte
   GetConstraint = ""
   bByte = InStr(1, sDescription, "DF_")
   If bByte > 0 Then
      GetConstraint = Mid$(sDescription, bByte, Len(sDescription))
      bByte = InStr(1, GetConstraint, "'")
      If bByte > 0 Then GetConstraint = Left$(GetConstraint, bByte - 1)
   Else
      bByte = InStr(1, sDescription, "DEFZERO")
      If bByte > 0 Then
         sSql = "sp_unbindefault '" & RdoCol.Fields(2) & "." & RdoCol.Fields(3) & "'"
         clsADOCon.ExecuteSQL sSql
      End If
   End If
   
   
End Function

'8/9/05 Close open sets

Public Sub FormUnload(Optional bDontShowForm As Byte)
   Dim iList As Integer
   Dim iResultSets As Integer
   On Error Resume Next
   MdiSect.lblBotPanel.Caption = MdiSect.Caption
'   If Forms.Count < 3 Then
'      iResultSets = clsADOCon.rdoResultsets.Count
'      For iList = iResultSets - 1 To 0 Step -1
'         clsADOCon.rdoResultsets(iList).Close
'      Next
'   End If
   If bDontShowForm = 0 Then
      Select Case cUR.CurrentGroup
         Case "Insp"
            zGr2Fain.Show
         Case "Ondc"
            zGr4Dock.Show
         Case "Rejt"
            zGr1Insp.Show
         Case "Stat"
            zGr3Stat.Show
         Case "Admn"
            zGr5Admn.Show
      End Select
      Erase bActiveTab
      cUR.CurrentGroup = ""
   End If
   
End Sub

Public Sub OpenFavorite(sSelected As String)
   CloseForms
   If LTrim$(sSelected) = "" Then Exit Sub
   MouseCursor 13
   On Error GoTo modErr1
   Select Case sSelected
      Case "New Inspection Report"
         cUR.CurrentGroup = "Rejt"
         InspRTe01a.Show
      Case "Revise An Inspection Report"
         cUR.CurrentGroup = "Rejt"
         InspRTe02a.Show
      Case "Discrepancy Codes"
         cUR.CurrentGroup = "Rejt"
         InspRTe03a.Show
      Case "Inspection Reports (Report)"
         cUR.CurrentGroup = "Rejt"
         InspRTp01a.Show
      Case "Open Corrective Action"
         cUR.CurrentGroup = "Rejt"
         InspRTp04a.Show
      Case "Customer Information"
         cUR.CurrentGroup = "Admn"
         AdmnQAe01a.Show
      Case "Inspectors"
         cUR.CurrentGroup = "Rejt"
         InspRTe06a.Show
      Case "Vendor Information"
         cUR.CurrentGroup = "Admn"
         AdmnQAe02a.Show
      Case "Parts"
         cUR.CurrentGroup = "Admn"
         InvcINe01a.Show
      Case "Divisions"
         cUR.CurrentGroup = "Admn"
         SadmSLe01a.Show
      Case "Customer List"
         cUR.CurrentGroup = "Admn"
         SaleSLp02a.Show
      Case "Customer Directory"
         cUR.CurrentGroup = "Admn"
         SaleSLp03a.Show
      Case "Vendor List"
         cUR.CurrentGroup = "Admn"
         PurcPRp02a.Show
      Case "Vendor Directory"
         cUR.CurrentGroup = "Admn"
         PurcPRp03a.Show
      Case "Parts By Part Number"
         cUR.CurrentGroup = "Admn"
         InvcINp01a.Show
      Case "Parts By Part Description"
         cUR.CurrentGroup = "Admn"
         InvcINp02a.Show
      Case "Divisions Report"
         cUR.CurrentGroup = "Admn"
         SadmSLp01a.Show
      Case "Shops"
         cUR.CurrentGroup = "Admn"
         CapaCPe02a.Show
      Case "Shop Information"
         cUR.CurrentGroup = "Admn"
         CapaCPp02a.Show
      Case "Responsibility Codes"
         cUR.CurrentGroup = "Rejt"
         InspRTe04a.Show
      Case "Disposition Codes"
         cUR.CurrentGroup = "Rejt"
         InspRTe05a.Show
      Case "Deactivate An Inspector"
         cUR.CurrentGroup = "Rejt"
         InspRTf03a.Show
      Case "Reactivate An Inspector"
         cUR.CurrentGroup = "Rejt"
         InspRTf04a.Show
      Case "Delete A Responsibility Code"
         cUR.CurrentGroup = "Rejt"
         InspRTf06a.Show
      Case "Delete A Disposition Code"
         cUR.CurrentGroup = "Rejt"
         InspRTf07a.Show
      Case "Delete A Discrepancy Code"
         cUR.CurrentGroup = "Rejt"
         InspRTf05a.Show
      Case "Change An Inspection Report Type Flag"
         cUR.CurrentGroup = "Rejt"
         InspRTf02a.Show
      Case "Delete An Inspection Report"
         cUR.CurrentGroup = "Rejt"
         InspRTf01a.Show
      Case "Inspections By Inspector"
         cUR.CurrentGroup = "Rejt"
         InspRTp08a.Show
      Case "Inspection Report Log"
         cUR.CurrentGroup = "Rejt"
         InspRTp02a.Show
      Case "Inspection Reports By Discrepancy Code"
         cUR.CurrentGroup = "Rejt"
         InspRTp03a.Show
      Case "Inspection Reports By Customer"
         cUR.CurrentGroup = "Rejt"
         InspRTp05a.Show
      Case "Inspection Reports By Vendor"
         cUR.CurrentGroup = "Rejt"
         InspRTp06a.Show
      Case "Inspection Reports By Division"
         cUR.CurrentGroup = "Rejt"
         InspRTp07a.Show
      Case "Inspection Report History By Part"
         cUR.CurrentGroup = "Rejt"
         InspRTp09a.Show
      Case "Part Family ID's"
         cUR.CurrentGroup = "Stat"
         StatSPe06a.Show
      Case "Family ID's"
         cUR.CurrentGroup = "Stat"
         StatSPe05a.Show
      Case "Process ID's"
         cUR.CurrentGroup = "Stat"
         StatSPe04a.Show
      Case "Team Members"
         cUR.CurrentGroup = "Stat"
         StatSPe02a.Show
      Case "Characteristic Reasoning Codes"
         cUR.CurrentGroup = "Stat"
         StatSPe03a.Show
      Case "List Of Team Members"
         cUR.CurrentGroup = "Stat"
         StatSPp01a.Show
      Case "List Of Reasoning Codes"
         cUR.CurrentGroup = "Stat"
         StatSPp02a.Show
      Case "List Of Process ID's"
         cUR.CurrentGroup = "Stat"
         StatSPp03a.Show
      Case "List Of Family ID's"
         cUR.CurrentGroup = "Stat"
         StatSPp04a.Show
      Case "SPC Processes"
         cUR.CurrentGroup = "Stat"
         StatSPe01a.Show
      Case "Delete A Reasoning Code"
         cUR.CurrentGroup = "Stat"
         StatSPf01a.Show
      Case "Delete A Process ID"
         cUR.CurrentGroup = "Stat"
         StatSPf02a.Show
      Case "Delete A Family ID"
         cUR.CurrentGroup = "Stat"
         StatSPf03a.Show
      Case "Update Family ID By Product Code"
         cUR.CurrentGroup = "Stat"
         StatSPf04a.Show
      Case "Processes By Part Number"
         cUR.CurrentGroup = "Stat"
         StatSPp05a.Show
      Case "Part Numbers By Family ID"
         cUR.CurrentGroup = "Stat"
         StatSPp06a.Show
      Case "Part Numbers By Process ID"
         cUR.CurrentGroup = "Stat"
         StatSPp07a.Show
      Case "Cause Codes"
         cUR.CurrentGroup = "Rejt"
         InspRTe07a.Show
      Case "Set On Dock Requirements"
         cUR.CurrentGroup = "Dock"
         DockODe03a.Show
      Case "On Dock Inspection"
         cUR.CurrentGroup = "Dock"
         DockODe01a.Show
      Case "On Dock (Delivered)"
         cUR.CurrentGroup = "Dock"
         DockODe02a.Show
      Case "Parts Requiring On Dock Inspection"
         cUR.CurrentGroup = "Dock"
         DockODp01a.Show
      Case "Parts On Dock (Delivered)"
         cUR.CurrentGroup = "Dock"
         DockODp02a.Show
      Case "Acceptance And Rejections By Vendor"
         cUR.CurrentGroup = "Dock"
         DockODp03a.Show
      Case "On Dock Inspections By Inspector"
         cUR.CurrentGroup = "Dock"
         DockODp04a.Show
      Case "Revise A First Article Report"
         cUR.CurrentGroup = "Insp"
         FainFAe02a.Show
      Case "New First Article Report"
         cUR.CurrentGroup = "Insp"
         FainFAe01a.Show
      Case "Copy A First Article Report"
         cUR.CurrentGroup = "Insp"
         FainFAf01a.Show
      Case "First Article Inspection Report"
         cUR.CurrentGroup = "Insp"
         FainFAp01a.Show
      Case "List Of First Article Reports"
         cUR.CurrentGroup = "Insp"
         FainFAp02a.Show
      Case "Cancel An On Dock Inspection"
         cUR.CurrentGroup = "Dock"
         DockODf01a.Show
      Case Else
         MouseCursor 0
   End Select
   Exit Sub
   
modErr1:
   On Error GoTo 0
   
End Sub







Sub Main()
   Dim sAppTitle$
   If App.PrevInstance Then
      On Error Resume Next
      sAppTitle = App.Title
      App.Title = "E1ePr"
      SysMsgBox.Width = 3800
      SysMsgBox.msg.Width = 3200
      SysMsgBox.tmr1.Enabled = True
      SysMsgBox.msg = sAppTitle & " Is Already Open."
      SysMsgBox.Show
      Sleep 5000
      AppActivate sAppTitle
   End
   Exit Sub
End If
' Set the Module name before loading the form
sProgName = "Quality Assurance"
MainLoad "qual"
GetFavorites "EsiQual"
' MM 9/10/2009
'sProgName = "Quality Assurance"
MdiSect.Show

End Sub



Public Sub UpdateTables()
   If MdiSect.bUnloading = 1 Then Exit Sub
   'Dim RdoTest As ADODB.Recordset
   Dim lRevision As Long
   
   SaveSetting "Esi2000", "AppTitle", "qual", "ESI Quality"
   SysOpen.Show
   SysOpen.prg1.Visible = True
   SysOpen.pnl = "Configuration Settings."
   SysOpen.pnl.Refresh
   
   MouseCursor 13
   On Error Resume Next
   'See OldUpdates 10/6/06
   '   2/6/07
   '   '5/30/06
   '    ConvertRejTagTables
   '    ConvertFAInspTables
   '    BuildKeys
   
   SysOpen.prg1.Value = 30
   Sleep 500
   SysOpen.prg1.Value = 70
   Sleep 500
   GoTo modErr2
   Exit Sub
   
modErr1:
   Resume modErr2
modErr2:
   'Set RdoTest = Nothing
   On Error GoTo 0
   SysOpen.Timer1.Enabled = True
   SysOpen.prg1.Value = 100
   SysOpen.Refresh
   Sleep 500
   
End Sub


Public Sub OldUpdates()
   
End Sub

'5/30/06
'11/21/06 Revised If Err

Public Sub BuildKeys()
   On Error Resume Next
   clsADOCon.ADOErrNum = 0
   clsADOCon.BeginTrans
   sSql = "DROP INDEX RjkyTable.KeyRef"
   clsADOCon.ExecuteSQL sSql
   If clsADOCon.ADOErrNum = 0 Then
      sSql = "ALTER TABLE RjkyTable ALTER COLUMN KEYREF CHAR(30) NOT NULL"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "ALTER TABLE RjkyTable ALTER COLUMN KEYDIM CHAR(15) NOT NULL"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "ALTER TABLE RjkyTable ADD Constraint PK_RjkyTable_KEYREF PRIMARY KEY CLUSTERED (KEYREF,KEYDIM) " _
             & "WITH FILLFACTOR=80 "
      clsADOCon.ExecuteSQL sSql
      If clsADOCon.ADOErrNum = 0 Then clsADOCon.CommitTrans Else clsADOCon.RollbackTrans
      
      Err.Clear
      clsADOCon.ADOErrNum = 0
      clsADOCon.BeginTrans
      sSql = "DROP INDEX RjdtTable.DatRef"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "ALTER TABLE RjdtTable ALTER COLUMN DATREF CHAR(30) NOT NULL"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "ALTER TABLE RjdtTable ALTER COLUMN DATKEY CHAR(15) NOT NULL"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "ALTER TABLE RjdtTable ADD Constraint PK_RjdtTable_DATREF PRIMARY KEY CLUSTERED (DATREF,DATKEY) " _
             & "WITH FILLFACTOR=80 "
      clsADOCon.ExecuteSQL sSql
      If clsADOCon.ADOErrNum = 0 Then clsADOCon.CommitTrans Else clsADOCon.RollbackTrans
      
      Err.Clear
      clsADOCon.ADOErrNum = 0
      clsADOCon.BeginTrans
      sSql = "DROP INDEX RinsTable.InsId"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "ALTER TABLE RinsTable ALTER COLUMN INSID CHAR(12) NOT NULL"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "ALTER TABLE RinsTable ADD Constraint PK_RinsTable_INSID PRIMARY KEY CLUSTERED (INSID) " _
             & "WITH FILLFACTOR=80 "
      clsADOCon.ExecuteSQL sSql
      If clsADOCon.ADOErrNum = 0 Then clsADOCon.CommitTrans Else clsADOCon.RollbackTrans
      
      Err.Clear
      clsADOCon.ADOErrNum = 0
      clsADOCon.BeginTrans
      sSql = "DROP INDEX RjfmTable.FamRef"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "ALTER TABLE RjfmTable ALTER COLUMN FAMREF CHAR(15) NOT NULL"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "ALTER TABLE RjfmTable ADD Constraint PK_RjfmTable_FAMREF PRIMARY KEY CLUSTERED (FAMREF) " _
             & "WITH FILLFACTOR=80 "
      clsADOCon.ExecuteSQL sSql
      If clsADOCon.ADOErrNum = 0 Then clsADOCon.CommitTrans Else clsADOCon.RollbackTrans
      
      Err.Clear
      clsADOCon.ADOErrNum = 0
      clsADOCon.BeginTrans
      sSql = "DROP INDEX RjcaTable.CauseRef"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "ALTER TABLE RjcaTable ALTER COLUMN CAUSEREF CHAR(12) NOT NULL"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "ALTER TABLE RjcaTable ADD Constraint PK_RjcaTable_CAUSEREF PRIMARY KEY CLUSTERED (CAUSEREF) " _
             & "WITH FILLFACTOR=80 "
      clsADOCon.ExecuteSQL sSql
      If clsADOCon.ADOErrNum = 0 Then clsADOCon.CommitTrans Else clsADOCon.RollbackTrans
      
      Err.Clear
      clsADOCon.ADOErrNum = 0
      clsADOCon.BeginTrans
      
      sSql = "DROP INDEX RjcdTable.CdeRef"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "ALTER TABLE RjcdTable ALTER COLUMN CDEREF CHAR(12) NOT NULL"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "ALTER TABLE RjcdTable ADD Constraint PK_RjcaTable_CDEREF PRIMARY KEY CLUSTERED (CDEREF) " _
             & "WITH FILLFACTOR=80 "
      clsADOCon.ExecuteSQL sSql
      If clsADOCon.ADOErrNum = 0 Then clsADOCon.CommitTrans Else clsADOCon.RollbackTrans
      
      Err.Clear
      clsADOCon.BeginTrans
      clsADOCon.ADOErrNum = 0
      sSql = "DROP INDEX RjdsTable.DisRef"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "ALTER TABLE RjdsTable ALTER COLUMN DISREF CHAR(12) NOT NULL"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "ALTER TABLE RjdsTable ADD Constraint PK_RjdsTable_DISREF PRIMARY KEY CLUSTERED (DISREF) " _
             & "WITH FILLFACTOR=80 "
      clsADOCon.ExecuteSQL sSql
      If clsADOCon.ADOErrNum = 0 Then clsADOCon.CommitTrans Else clsADOCon.RollbackTrans
      
      Err.Clear
      clsADOCon.ADOErrNum = 0
      clsADOCon.BeginTrans
      sSql = "DROP INDEX RjktTable.TeamRef"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "ALTER TABLE RjktTable ALTER COLUMN TEAMREF CHAR(15) NOT NULL"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "ALTER TABLE RjktTable ALTER COLUMN TEAMDIM CHAR(15) NOT NULL"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "ALTER TABLE RjktTable ALTER COLUMN TEAMID CHAR(15) NOT NULL"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "ALTER TABLE RjktTable ADD Constraint PK_RjktTable_KEYREF PRIMARY KEY CLUSTERED (TEAMREF,TEAMDIM,TEAMID) " _
             & "WITH FILLFACTOR=80 "
      clsADOCon.ExecuteSQL sSql
      If clsADOCon.ADOErrNum = 0 Then clsADOCon.CommitTrans Else clsADOCon.RollbackTrans
      
      Err.Clear
      clsADOCon.ADOErrNum = 0
      clsADOCon.BeginTrans
      sSql = "DROP INDEX RjmmTable.MemRef"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "ALTER TABLE RjmmTable ALTER COLUMN MEMREF CHAR(30) NOT NULL"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "ALTER TABLE RjmmTable ALTER COLUMN MEMKEY CHAR(15) NOT NULL"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "ALTER TABLE RjmmTable ALTER COLUMN MEMID CHAR(15) NOT NULL"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "ALTER TABLE RjmmTable ADD Constraint PK_RjmmTable_MEMREF PRIMARY KEY CLUSTERED (MEMREF,MEMKEY,MEMID) " _
             & "WITH FILLFACTOR=80 "
      clsADOCon.ExecuteSQL sSql
      If clsADOCon.ADOErrNum = 0 Then clsADOCon.CommitTrans Else clsADOCon.RollbackTrans
      
      Err.Clear
      clsADOCon.ADOErrNum = 0
      clsADOCon.BeginTrans
      sSql = "DROP INDEX RjrcTable.RcoRef"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "ALTER TABLE RjrcTable ALTER COLUMN RCOREF CHAR(15) NOT NULL"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "ALTER TABLE RjrcTable ADD Constraint PK_RjrcTable_RCOREF PRIMARY KEY CLUSTERED (RCOREF) " _
             & "WITH FILLFACTOR=80 "
      clsADOCon.ExecuteSQL sSql
      If clsADOCon.ADOErrNum = 0 Then clsADOCon.CommitTrans Else clsADOCon.RollbackTrans
      
      Err.Clear
      clsADOCon.ADOErrNum = 0
      clsADOCon.BeginTrans
      sSql = "DROP INDEX RjrsTable.ResRef"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "ALTER TABLE RjrsTable ALTER COLUMN RESREF CHAR(12) NOT NULL"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "ALTER TABLE RjrsTable ADD Constraint PK_RjrsTable_RESREF PRIMARY KEY CLUSTERED (RESREF) " _
             & "WITH FILLFACTOR=80 "
      clsADOCon.ExecuteSQL sSql
      If clsADOCon.ADOErrNum = 0 Then clsADOCon.CommitTrans Else clsADOCon.RollbackTrans
      
      Err.Clear
      clsADOCon.ADOErrNum = 0
      clsADOCon.BeginTrans
      
      sSql = "DROP INDEX RjitTable.RitRef"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "ALTER TABLE RjitTable ALTER COLUMN RITREF CHAR(12) NOT NULL"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "ALTER TABLE RjitTable ALTER COLUMN RITITM SMALLINT NOT NULL"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "ALTER TABLE RjitTable ADD Constraint PK_RjitTable_RITREF PRIMARY KEY CLUSTERED (RITREF,RITITM) " _
             & "WITH FILLFACTOR=80 "
      clsADOCon.ExecuteSQL sSql
      If clsADOCon.ADOErrNum = 0 Then clsADOCon.CommitTrans Else clsADOCon.RollbackTrans
      
      Err.Clear
      clsADOCon.ADOErrNum = 0
      clsADOCon.BeginTrans
      
      sSql = "DROP INDEX RjtmTable.TmmRef"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "ALTER TABLE RjtmTable ALTER COLUMN TMMID CHAR(15) NOT NULL"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "ALTER TABLE RjtmTable ADD Constraint PK_RjtmTable_TMMREF PRIMARY KEY CLUSTERED (TMMID) " _
             & "WITH FILLFACTOR=80 "
      clsADOCon.ExecuteSQL sSql
      If clsADOCon.ADOErrNum = 0 Then clsADOCon.CommitTrans Else clsADOCon.RollbackTrans
   Else
      clsADOCon.RollbackTrans
   End If
   Exit Sub
   
KeysErr1:
   On Error Resume Next
   clsADOCon.RollbackTrans
   
End Sub

'5/30/06 Includes Keys

Public Sub ConvertFAInspTables()
   Dim bBadCol As Byte
   Dim sconstraint As String
   On Error Resume Next
   Err.Clear
   'Start FahdTable
   'FA_DRAWINGTOLPLUS
   sSql = "sp_columns @table_name=FahdTable,@column_name=FA_DRAWINGTOLPLUS"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then 'See Else
               sSql = "ALTER TABLE FahdTable ALTER COLUMN FA_DRAWINGTOLPLUS dec(12,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE FahdTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE FahdTable ALTER COLUMN FA_DRAWINGTOLPLUS DEC(12,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'FahdTable.FA_DRAWINGTOLPLUS'"
               clsADOCon.ExecuteSQL sSql
            Else
               GoTo Keys
            End If
         End If
      End If
   End With
   'FA_DRAWINGTOLMINUS
   sSql = "sp_columns @table_name=FahdTable,@column_name=FA_DRAWINGTOLMINUS"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE FahdTable ALTER COLUMN FA_DRAWINGTOLMINUS dec(12,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE FahdTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE FahdTable ALTER COLUMN FA_DRAWINGTOLMINUS DEC(12,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'FahdTable.FA_DRAWINGTOLMINUS'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'FA_ANGLETOLPLUS
   sSql = "sp_columns @table_name=FahdTable,@column_name=FA_ANGLETOLPLUS"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE FahdTable ALTER COLUMN FA_ANGLETOLPLUS dec(12,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE FahdTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE FahdTable ALTER COLUMN FA_ANGLETOLPLUS DEC(12,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'FahdTable.FA_ANGLETOLPLUS'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   
   'FA_ANGLETOLMINUS
   sSql = "sp_columns @table_name=FahdTable,@column_name=FA_ANGLETOLMINUS"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE FahdTable ALTER COLUMN FA_ANGLETOLMINUS dec(12,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE FahdTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE FahdTable ALTER COLUMN FA_ANGLETOLMINUS DEC(12,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'FahdTable.FA_ANGLETOLMINUS'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'No FaitTable Real values
   
Keys:
   Err.Clear
   clsADOCon.ADOErrNum = 0
   
   sSql = "DROP INDEX FahdTable.FahdRef"
   clsADOCon.ExecuteSQL sSql
   If clsADOCon.ADOErrNum = 0 Then
      sSql = "ALTER TABLE FahdTable ALTER COLUMN FA_REF CHAR(30) NOT NULL"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "ALTER TABLE FahdTable ALTER COLUMN FA_REVISION CHAR(6) NOT NULL"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "ALTER TABLE FahdTable ADD Constraint PK_FahdTable_FAREF PRIMARY KEY CLUSTERED (FA_REF,FA_REVISION) " _
             & "WITH FILLFACTOR=80 "
      clsADOCon.ExecuteSQL sSql
      
      sSql = "ALTER TABLE FaitTable ADD CONSTRAINT FK_FaitTable_FahdTable FOREIGN KEY (FA_ITNUMBER,FA_ITREVISION) References FahdTable"
      clsADOCon.ExecuteSQL sSql
   End If
   
End Sub
