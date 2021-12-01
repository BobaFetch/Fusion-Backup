Attribute VB_Name = "Errors"
Option Explicit

Public Sub ProcessError(SubName As String)
   MouseCursor ccDefault
   
   If (Not clsADOCon Is Nothing) Then
      clsADOCon.RollbackTrans
   End If
   
   sProcName = SubName
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   Dim frm As Form
   Set frm = MdiSect.ActiveForm
   If frm Is Nothing Then
      Set frm = New ClassErrorForm
   End If
   DoModuleErrors MdiSect.ActiveForm
End Sub

Public Sub DoModuleErrors(frm As Form)
   Dim RdoEvent As ADODB.Recordset
   Dim bByte As Byte
   Dim bInstr As Byte
   '   Dim iFreeFile    As Integer
   Dim iWarningType As Integer
   Dim sMsg As String
   Dim sMsg2 As String
   
   'error log
   Dim sDate As String * 16
   Dim sSection As String * 8
   Dim sForm As String * 12
   Dim sErrNum As String * 10
   Dim sErrSev As String * 2
   Dim sProc As String * 10
   Dim sUserName As String * 20
   Dim ErringSQL As String
   
   ErringSQL = sSql        'save sql - it gets overwritten below
   MouseCursor 13
   LockWindowUpdate 0
   On Error Resume Next
   sDate = Format(GetServerDateTime, "mm/dd/yy hh:mm")
   sSection = Left(sProgName, 5)
   sForm = Trim(frm.Name)
   sErrNum = str(CurrError.Number)
   sUserName = Left(cUR.CurrentUser, 18)
   
   MsgBox "ERROR:" & CStr(CurrError.Number)
   MsgBox "ERROR:" & CStr(CurrError.Description)

   
   'Default Warning Flag. Setting bByte to True changes
   'the Warning Flag and smooth closes the app if req'd.
   iWarningType = vbExclamation
   Select Case CurrError.Number
      Case 3
         sMsg = "Return Without GoSub"
         bByte = 0
      Case 5
         sMsg = "Invalid Procedure Call"
         bByte = 0
      Case 6
         sMsg = "Overflow"
         bByte = 0
      Case 7
         sMsg = "Out Of Memory"
         bByte = 1
      Case 9
         sMsg = "Subscript Out Of Range"
         bByte = 0
      Case 10
         sMsg = "This Array Is Fixed Or Temporarily Locked"
         bByte = 1
      Case 11
         sMsg = "Division By Zero"
         bByte = 1
      Case 13
         sMsg = "Type Mismatch"
         bByte = 0
      Case 14
         sMsg = "Out Of String Space"
         bByte = 0
      Case 16
         sMsg = "Expression Too Complex"
         bByte = 0
      Case 17
         sMsg = "Can't Perform Requested Operation"
         bByte = 0
      Case 18
         sMsg = "User Interrupt Occurred"
         bByte = 1
      Case 20
         sMsg = "Resume Without Error"
         bByte = 0
      Case 28
         sMsg = "Out Of Strack Space"
         bByte = 1
      Case 35
         sMsg = "Sub, Function, Or Property Not Defined"
         bByte = 0
      Case 47
         sMsg = "Too Many DLL Application Clients"
         bByte = 1
      Case 48
         sMsg = "Error In Loading DLL"
         bByte = 1
      Case 49
         sMsg = "Bad DLL Calling Convention"
         bByte = 1
      Case 51
         sMsg = "Internal Error"
         bByte = 1
      Case 52
         sMsg = "Bad File Name Or Number"
         bByte = 1
      Case 53
         sMsg = "File Not Found"
         bByte = 0
         iWarningType = vbInformation
      Case 54
         sMsg = "Bad File Mode"
         bByte = 0
      Case 55
         sMsg = "File Already Open"
         bByte = 0
      Case 57
         sMsg = "Device I/O Error"
         bByte = 0
      Case 58
         sMsg = "File Already Exists"
         bByte = 0
      Case 59
         sMsg = "Bad Record Length"
         bByte = 1
      Case 61
         sMsg = "Disk Full"
         bByte = 1
      Case 62
         sMsg = "Input Past End Of File"
         bByte = 0
      Case 63
         sMsg = "Bad Record Number"
         bByte = 0
      Case 67
         sMsg = "Too Many Files"
         bByte = 1
      Case 68
         sMsg = "Device Unavailable"
         bByte = 0
      Case 70
         sMsg = "Permission Denied"
         bByte = 0
      Case 71
         sMsg = "Disk Not Ready"
         bByte = 0
      Case 74
         sMsg = "Can't Rename With Different Drive"
         bByte = 0
      Case 75
         sMsg = "Path/File Access Error"
         bByte = 0
      Case 76
         sMsg = "Path Not Found." & vbCrLf _
                & "An Attempt Was Made To Open A File From An" & vbCrLf _
                & "Invalid Location. Check The User Filepath Settings."
         bByte = 0
      Case 91
         sMsg = "Object Variable Or With Block Variable Not Set."
         sMsg = sMsg & vbCrLf & "Check Network Connection"
         bByte = 0
      Case 94
         sMsg = "Invalid Use Of Null. Please Report This Error."
         bByte = 0
      Case 340
         sMsg = "Control Element Does Not Exist."
         bByte = 0
      Case 380
         sMsg = "Invalid Property Value"
         bByte = 0
      Case 438
         sMsg = "Function Not Supported By Object.  Please Report This Error."
         bByte = 0
      Case 482, 483, 486
         sMsg = "Printer Error"
         bByte = 0
         'Jet
      Case 3001 To 3648
         sMsg = "JET DSS Database Error. Contact Systems Administrator."
         bByte = 0
         'Crystal
      Case 20500
         sMsg = "Not Enough Memory To Complete Report. " & vbCrLf _
                & "SSCSDK32.DLL May Be Missing Or Corrupt."
         bByte = 0
      Case 20501 To 20506
         sMsg = "Undocumented Crystal Reports Error. " & vbCrLf _
                & "Contact Systems Administrator."
         CurrError.Description = Left(sMsg, 34)
         bByte = 0
      Case 20507
         sMsg = "Report Wasn't Found Or Couldn't Be Loaded. " & vbCrLf _
                & "Check Your Report Path In Settings."
         CurrError.Description = "Couldn't Find The Requested Report."
         bByte = 0
      Case 20508, 20509, 20511, 20512, 20514
         sMsg = "Undocumented Crystal Reports Error. " & vbCrLf _
                & "Contact Systems Administrator."
         CurrError.Description = Left(sMsg, 34)
         bByte = 0
      Case 20510
         'Crystal does not broadcast the name
         sMsg = "The Requested Formula Doesn't Exist (Invalid Formula Name). "
         CurrError.Description = Left(sMsg, 35)
         bByte = 0
      Case 20513
         sMsg = "The Requested Printer Is Not Valid. " & vbCrLf _
                & "Contact Systems Administrator."
         CurrError.Description = Left(sMsg, 34)
         bByte = 0
      Case 20515
         bInstr = InStr(1, CurrError.Description, "<Record_Selection>")
         If bInstr > 0 Then
            CurrError.Description = "Error In The SQL Query (Record Selection)"
            sMsg = "Error In SQL Statement  (Record Selection)" & vbCrLf _
                   & "Contact Your Systems Administrator."
         Else
            bInstr = InStr(1, CurrError.Description, "<")
            CurrError.Description = "Error In Formula " & Mid(CurrError.Description, bInstr)
            sMsg = CurrError.Description & vbCrLf _
                   & "Contact Your Systems Administrator."
         End If
         bByte = 0
      Case 20516, 20517, 20519
         sMsg = "Not Enough Windows Resources To Complete Report. " & vbCrLf _
                & "Close Some Applications,  " & sSysCaption & " And Restart."
         bByte = 0
      Case 20518
         sMsg = "An Attempt Was Made To Hide Or Show A Report Group" & vbCrLf _
                & "That Does Not Exist. Contact Your Systems Administrator."
         CurrError.Description = "Hide/Show A Group Does Not Exist"
         bByte = 0
      Case 20520
         sMsg = "Print Job Started And Report In Progress Or " & vbCrLf _
                & "There Is No Default Print Or The Printer Is Offline." & vbCrLf _
                & "Crystal Reports Notice Not An Error."
         iWarningType = vbInformation
         bByte = 0
      Case 20521 To 20522
         sMsg = "Not Enough Windows Resources To Complete Report. " & vbCrLf _
                & "Close Some Applications, " & sSysCaption & " And Restart."
         bByte = 0
      Case 20523
         sMsg = "Invalid Character In The GroupConditioning." & vbCrLf _
                & "Contact Your Systems Administrator."
         CurrError.Description = Left(sMsg, 35)
         bByte = 0
      Case 20524
         sMsg = "Undocumented Crystal Reports Error. " & vbCrLf _
                & "Contact Systems Administrator."
         CurrError.Description = Left(sMsg, 34)
      Case 20525
         sMsg = "Report Is Damaged. Unable To Open Report. " & vbCrLf _
                & "Contact Your Systems Administrator."
         bByte = 0
      Case 20526
         sMsg = "No Default Printer Has Been Set. "
         CurrError.Description = sMsg
         bByte = 0
      Case 20527
         sMsg = "Error In SQL Server Connection. " & vbCrLf _
                & "The Report Cannot Resolve The Data."
         CurrError.Description = "Crystal Reports Cannot Resolve The Data"
         bByte = 0
      Case 20529
         sMsg = "Your Disk Drive Is Full And Files May Be Lost." & vbCrLf _
                & "Exit " & sSysCaption & " And Free Resources."
         bByte = 0
      Case 20530
         sMsg = "An Attempt Was Made To Access A Report In Use. " & vbCrLf _
                & "Try The Report Again."
         CurrError.Description = "An Attempt Was Made To Access A Report In Use"
         iWarningType = vbInformation
         bByte = 0
      Case 20531
         sMsg = "Incorrect Password. Permission Denied."
         bByte = 0
      Case 20532
         sMsg = "File I/O Error. Disk Problem Other Than Full." & vbCrLf _
                & "Exit " & sSysCaption & " Contact Systems Administrator."
         bByte = 0
      Case 20533
         sMsg = "Unable To Open The Database File." & vbCrLf _
                & "Contact Your Systems Administrator."
         bByte = 0
      Case 20534
         sMsg = "The Database Columns Are Not Correctly Mapped" & vbCrLf _
                & "Or Have A Collation Conflict. Contact Your" & vbCrLf _
                & "Systems Administrator."
         CurrError.Description = "Column Mapping Or Collation Problem."
         bByte = 0
      Case 20535 To 20543
         sMsg = CurrError.Description
         bByte = 0
      Case 20544
         sMsg = "This Report Is Open By Another User." & vbCrLf _
                & "Try The Report Again In A Few Minutes."
         bByte = 0
      Case 20545
         sMsg = CurrError.Description
         iWarningType = vbInformation
         bByte = 0
      Case 20546 To 20598
         sMsg = CurrError.Description
         bByte = 0
      Case 20599
         sMsg = "ODBC Permissions/Access Error. Check ODBC Data Source." & vbCrLf _
                & "DSN " & sDsn & " May Be Improperly Installed Or Does Not Exist."
         bByte = 0
      Case 20600 To 20996
         sMsg = "Undocumented Crystal Reports Error." & vbCrLf _
                & "Contact Your Systems Administrator."
      Case 20997
         sMsg = "Invalid Report Path Or No Network Permissions." & vbCrLf _
                & "Check Your Report Path And Server Permissions."
         bByte = 0
      Case 20998
         sMsg = "Report Path Is Too Long.     " & vbCrLf _
                & "Use A Mapped Path Name Instead (x:\somedir) Or " _
                & "Possible Mismatch Of Graph DLL Libraries."
         bByte = 0
         'Rdo
      Case 40000
         sMsg = "An Error Occurred Configuring The DataSource Name."
         bByte = 0
      Case 40001
         sMsg = "SQL Returned No Data Found From Query."
         iWarningType = vbInformation
         bByte = 0
      Case 40002
         If Left(CurrError.Description, 5) = "22003" Then
            sMsg = Left(CurrError.Description, 5) & "-Attempted Insert A Value Greater" & vbCrLf _
                   & "Than The Requested Column Allows."
            bByte = 0
         Else
            If Left(CurrError.Description, 5) = "S0002" Then
               sMsg = Left(CurrError.Description, 5) & "-The Requested Table Wasn't Found."
               bByte = 0
            Else
               If Left(CurrError.Description, 5) = "S0022" Then
                  'sMsg = Left(CurrError.Description, 5) & "-The Requested Column Wasn't Found."
                  sMsg = CurrError.Description
                  bByte = 0
               Else
                  If Left(CurrError.Description, 5) = "01000" Then
                     'sMsg = "An Attempt Was Made To Add A Duplicate Record" & vbCrLf _
                     '       & "Or The Data To Be Inserted Is Not Valid."
                     sMsg = CurrError.Description
                     bByte = 0
                  Else
                     If InStr(CurrError.Description, "0851") > 0 Then
                        sMsg = "ODBC Link Was Lost. Reconnection Required."
                        bByte = 1
                     ElseIf InStr(CurrError.Description, "S1T00") Then _
                                     sMsg = "The Query Has Timed Out, Check Settings."
                        bByte = 0
                     ElseIf InStr(CurrError.Description, "S1010") Then _
                                     sMsg = "The Query Has Timed Out. Bulk Copy May Be Turned On/"
                        bByte = 0
                     ElseIf InStr(CurrError.Description, "08S01") Then _
                                     sMsg = "The Network Connection With SQL Server Failed."
                        bByte = 0
                     ElseIf InStr(CurrError.Description, "07S01") Then _
                                     sMsg = "Invalid Default Parameter In Query Results."
                        bByte = 0
                        Err.Clear
                        Exit Sub
                     Else
                        sMsg = "Internal ODBC Error Encountered."
                        bByte = 0
                     End If
                  End If
               End If
            End If
            If Left(CurrError.Description, 5) = "37000" Then
               'Changed some for SSL 7.0
'               sMsg = Left(CurrError.Description, 5) & vbCrLf _
'                      & "The Cursor Is No Longer Open. Invalid Character " & vbCrLf _
'                      & "Found Or The Query Couldn't Process Requested Data. " & vbCrLf _
'                      & "Please Report This To Your Systems Administrator."
               sMsg = Left(CurrError.Description, 5) & vbCrLf _
                      & "SQL error. Invalid Character: " & vbCrLf _
                      & CurrError.Description & vbCrLf _
                      & "Please Report This To Your Systems Administrator."
               bByte = 0
            End If
         End If
      Case 40003
         sMsg = "An Invalid Value For The Cursor Driver Was Passed."
         bByte = 0
      Case 40004
         sMsg = "An Invalid ODBC Handle Was Encountered."
         bByte = 0
      Case 40005
         sMsg = "Invalid Connection String."
         bByte = 1
      Case 40006
         sMsg = "An Unexpected Error Occurred."
         bByte = 0
      Case 40008
         sMsg = "Invalid Operation For Forward-Only Cursor."
         bByte = 0
      Case 40009
         sMsg = "No Current Row (No Matching Query Data Found)."
         iWarningType = vbInformation
         bByte = 0
      Case 40010
         sMsg = "Invalid Row For Add New."
         bByte = 0
      Case 40011
         sMsg = "Object Is Invalid Or Not Set."
         bByte = 0
      Case 40012
         sMsg = "Invalid Seek Flag."
         bByte = 0
      Case 40013
         sMsg = "Partial Equality Requires String Column."
         bByte = 0
      Case 40014
         sMsg = "Incompatible Data Types For Compare."
         bByte = 0
      Case 40015
         sMsg = "Can't Create Prepared Statement."
         bByte = 0
      Case 40016
         sMsg = "Version.DLL Error."
         bByte = 1
      Case 40017, 40018
         sMsg = "Can't Execute Statement."
         bByte = 0
      Case 40019
         sMsg = "An Invalid Value For The Concurrency Option."
         bByte = 0
      Case 40020
         sMsg = "Can't Open Result Set For Unnamed Table."
         bByte = 0
      Case 40021
         sMsg = "Object Collection Error."
         bByte = 0
      Case 40022
         sMsg = "The RDO Results Set Is Empty (No Data)."
         iWarningType = vbInformation
         bByte = 0
      Case 40023
         sMsg = "Invalid State For Cursor Move. "
         bByte = 0
      Case 40024
         sMsg = "Already Beyond The End Of The Result Set."
         bByte = 0
      Case 40025
         sMsg = "BOF Already Set."
         bByte = 0
      Case 40026
         sMsg = "Invalid Result Set State For Update."
         bByte = 0
      Case 40027
         sMsg = "Invalid Bookmark Or No Bookmark Allowed."
         bByte = 0
      Case 40028
         sMsg = "Invalid Bookmark Argument To Move."
         bByte = 0
      Case 40029
         sMsg = "Current Row As EOF/BOF Already Set."
         bByte = 0
      Case 40030
         sMsg = "Already At BOF."
         bByte = 0
      Case 40031
         sMsg = "Already At EOF."
         bByte = 0
      Case 40032
         sMsg = "Couldn't Load The ODBC Installation Library."
         bByte = 1
      Case 40033
         sMsg = "An Invalid Value For The Prompt Option Was Passed."
         bByte = 1
      Case 40034
         sMsg = "An Invalid Value For The Cursor Type Parameter Was Passed."
         bByte = 0
      Case 40035
         sMsg = "Column Not Bound Correctly."
         bByte = 0
      Case 40036
         sMsg = "Unbound Column-Use Get Chunk Method."
         bByte = 0
      Case 40037
         sMsg = "Can't Assign Value To Unbound Column."
         bByte = 0
      Case 40038
         sMsg = "Can't Assign Value To Non-Updatable Field."
         bByte = 0
      Case 40039
         sMsg = "Can't Assign Value To Column Unless In Edit Mode."
         bByte = 0
      Case 40040
         sMsg = "Incorrect Type For Parameter."
         bByte = 0
      Case 40041
         sMsg = "Object Collection: Couldn't Find Column Requested By Query."
         bByte = 0
      Case 40042
         sMsg = "Can't Assign Value To Unbound Parameter."
         bByte = 0
      Case 40043
         sMsg = "Can't Assign Value To Output-Only Parameter."
         bByte = 0
      Case 40044
         sMsg = "Incorrect RDO Parameter Type."
         bByte = 0
      Case 40045
         sMsg = "Tried To Execute A Query With An Asynchronous Query In Progress."
         bByte = 0
      Case 40046
         sMsg = "The Object Has Already Been Closed."
         bByte = 0
      Case 40047
         sMsg = "Invalid Name For The Environment."
         bByte = 0
      Case 40048
         sMsg = "Environment Name Already Exists In The Collection."
         bByte = 0
      Case 40049
         sMsg = "Object Collection Is Read-Only."
         bByte = 0
      Case 40050
         sMsg = "Get New Enum: Couldn't Get Interface."
         bByte = 0
      Case 40051
         sMsg = "Assignment To Count Property Not Allowed."
         bByte = 0
      Case 40052
         sMsg = "You Must Use Append Chunk To Set Data In A Text Or Image."
         bByte = 0
      Case 40053
         sMsg = "Object Collection: Can't Add Non Object Item."
         bByte = 0
      Case 40054
         sMsg = "An Invalid Parameter Was Passed."
         bByte = 0
      Case 40055
         sMsg = "Invalid Operation."
         bByte = 0
      Case 40056
         sMsg = "The Row Has Been Deleted."
         bByte = 0
      Case 40057
         sMsg = "An Attempt Was Made To Issue A Select Statement Using Execute."
         bByte = 0
      Case 40058
         sMsg = "Can't Update Column, The Result Set Is Read Only."
         bByte = 0
      Case 40059
         sMsg = "Cancel Has Been Selected In An ODBC Dialog Requesting Parameters."
         iWarningType = vbInformation
         bByte = 0
      Case 40060
         sMsg = "Needs Chunk Required Flags."
         bByte = 0
      Case 40061
         sMsg = "Could Not Load Resource Library."
         bByte = 1
      Case 40069
         sMsg = "General Client Cursor Error."
         bByte = 0
      Case 40071
         sMsg = "The RDO Connection Object Is Not Connected To A Data Source."
         bByte = 1
      Case 40072
         sMsg = "The RDO Connection Object Is Already Connected To The Data Source."
         bByte = 0
      Case 40073
         sMsg = "The RDO Connection Object Is Busy Connecting " & vbCrLf _
                & "To The Data Source. Retry The Selection."
         bByte = 0
      Case 40074
         sMsg = "The RDO Query Or RDO Results Set Has No Active Connection Source."
         bByte = 1
      Case 40075
         sMsg = "Incorrect Cursor Driver."
         bByte = 0
      Case 40076
         sMsg = "This Property Is Currently Read Only."
         iWarningType = vbInformation
         bByte = 0
      Case 40077
         sMsg = "The Object Is Already In The Collection."
         iWarningType = vbInformation
         bByte = 0
      Case 40078
         sMsg = "Failed To Load RDOCURS.DLL"
         bByte = 1
      Case 40079
         sMsg = "Can't Find The Requested Table To Update."
         bByte = 0
      Case 40080, 40081, 40082, 40083, 40085
         sMsg = "Invalid RDO/SQL Server Option."
         bByte = 0
      Case 40088
         sMsg = "No Open Cursor For Transaction Commit."
         bByte = 0
      Case 40500, 40501, 40502, 40503
         sMsg = "Unexpected Internal RDO Error "
         bByte = 1
      Case 40504
         sMsg = "Could Not Refresh Controls."
         bByte = 0
      Case 40505
         sMsg = "Invalid Property Value."
         bByte = 0
      Case 40506
         sMsg = "Invalid Collection Object."
         bByte = 0
      Case 40507
         sMsg = "Method Cannot Be Called In RDO's Current State."
         bByte = 0
      Case 40508
         sMsg = "One Or More Of The Arguments Is Invalid."
         bByte = 0
      Case 40509
         sMsg = "Result Set Is Empty."
         iWarningType = vbInformation
         bByte = 0
      Case 40510
         sMsg = "Out Of Memory. Close " & sSysCaption & "."
         bByte = 1
      Case 40511
         sMsg = "Result Set Not Available."
         bByte = 0
      Case 40512
         sMsg = "The Connection Is Not Open."
         bByte = 1
      Case 40513, 40514
         sMsg = "Property Cannot Be Set In RDC's Current State."
         bByte = 0
      Case 40515
         sMsg = "Type Mismatch."
         bByte = 0
      Case 40516
         sMsg = "Cannot Connect To Remote Data Object."
         bByte = 1
      Case Else
'         sMsg = "Undocumented Error           "
         sMsg = CurrError.Description
         bByte = 1
   End Select
   
   'add Crystal Reports filename, error description,
   'formulas and SQL to message.
   Select Case CurrError.Number
      Case 20500 To 20999
         ' MM not need this is old crw.
         'sMsg2 = CrystalParameterString(frm)
         
         'get error description and replace lone cr's and lf's with crlf's
         Dim sCrystal As String
         sCrystal = Replace(CurrError.Description, vbCrLf, "!$")
         sCrystal = Replace(sCrystal, vbCrLf, "!$")
         sCrystal = Replace(sCrystal, vbLf, "!$")
         sCrystal = Replace(sCrystal, "!$", vbCrLf)
         CurrError.Description = sCrystal
         sMsg2 = sMsg2 & "Crystal Reports error description: " & sCrystal & vbCrLf
         
         'if Sql Server error, add SQL to log
      Case 40000 To 40516
         sMsg2 = "SQL Server Command: " & ErringSQL
         
      Case Else
         sMsg2 = sSysCaption & " System."
   End Select
   If iWarningType = vbInformation Then
      sMsg = "Notification" & str(CurrError.Number) & vbCrLf & sMsg & vbCrLf & sMsg2
   Else
      If bByte = 1 Then
         sMsg = "Error" & str(CurrError.Number) & vbCrLf & sMsg & vbCrLf & sMsg2
      Else
         sMsg = "Warning" & str(CurrError.Number) & vbCrLf & sMsg & vbCrLf & sMsg2
      End If
   End If
   MouseCursor 0
   'Show the user and do as required
   MdiSect.Enabled = True
   
   If bByte = 1 Then
      sErrSev = "16"
      '            If Len(Trim(CurrError.Description)) < 35 Then
      '                Print #iFreeFile, sDate; sSection; sForm; sUserName; _
      '                    sErrNum; sErrSev; " "; sProcName; " "; Trim(CurrError.Description)
      '            Else
      '                Print #iFreeFile, sDate; sSection; sForm; sUserName; _
      '                    sErrNum; sErrSev; " "; sProcName; " "; Left$(CurrError.Description, 35)
      '            End If
      '            Close iFreeFile
      sMsg = sMsg & vbCrLf & sProcName
      sMsg = sMsg & vbCrLf & "Contact System Administrator"
      MsgBox sMsg, vbCritical, frm.Caption
   Else
      If iWarningType = vbInformation Then sErrSev = "64" Else sErrSev = "48"
      sMsg = sMsg & vbCrLf & "Procedure: " & sProcName
      MsgBox sMsg, iWarningType, frm.Caption
      sErrSev = "48"
      CurrError.Number = 0
   End If
   
   Close
   Err.Clear
   
   'log the error in the SystemEvents table
   'this must happen on a separate connection, otherwise, when the transaction is rolled back,
   'the new SystemEvents row will be rolled back as well
   
   Dim clsADOCon2 As ClassFusionADO
   Dim strConStr As String
   Dim strSaAdmin As String
   Dim strSaPassword As String
   Dim strServer As String
   Dim strDBName As String
   Dim errnum    As Long
   Dim errdesc   As String
   
   ' DNS strServer = UCase(GetUserSetting(USERSETTING_ServerName))
   strServer = UCase(GetConfUserSetting(USERSETTING_ServerName))
   strSaAdmin = Trim(GetSysLogon(True))
   strSaPassword = Trim(GetSysLogon(False))
   strDBName = sDataBase
   
   Set clsADOCon2 = New ClassFusionADO
   
   strConStr = "Driver={SQL Server};Provider='sqloledb';UID=" & strSaAdmin & ";PWD=" & _
            strSaPassword & ";SERVER=" & strServer & ";DATABASE=" & strDBName & ";"
   
   If clsADOCon2.OpenConnection(strConStr, errnum, errdesc) = False Then
     MsgBox "An error occured while trying to do second connect to the specified database:" & Chr(13) & Chr(13) & _
            "Error Number = " & CStr(errnum) & Chr(13) & _
            "Error Description = " & errdesc, vbOKOnly + vbExclamation, "  DB Connection Error"
     GoTo CleanUp
   End If
   
   sSql = "SELECT * FROM SystemEvents"
   bSqlRows = clsADOCon2.GetDataSet(sSql, RdoEvent, ES_KEYSET)

   With RdoEvent
      .AddNew
      !Event_Date = Format$(Now, "mm/dd/yy h:mm AM/PM")
      !Event_Section = sSection
      !Event_Form = sForm
      !Event_User = sUserName
      !Event_Event = Val(sErrNum)
      !Event_Warning = Val(sErrSev)
      !Event_Procedure = sProcName
      !Event_Text = Left(Trim(sMsg), 4096)
      !Event_SQL = "Most Recent SQL: " & ErringSQL
      .Update
   End With
   Set RdoEvent = Nothing
   Set clsADOCon2 = Nothing
   
   If sErrSev = "16" Then
      CloseFiles
   End
Else
   sProcName = ""
End If

   Exit Sub

CleanUp:
 'clsADOCon.CleanupRecordset RS
 Set clsADOCon2 = Nothing

End Sub


