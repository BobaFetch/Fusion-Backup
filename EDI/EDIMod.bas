Attribute VB_Name = "EDIMod"
Option Explicit

Public bFoundPart As Byte

Public sCurrForm As String
Public sFavorites(13) As String
Public sRecent(10) As String
Public sSession(50) As String



''from ESIPROJ in other modules
Public sFilePath As String
'
'' RDO
'Public RdoCon As rdoConnection
'Public RdoEnv As rdoEnvironment
'Public RdoErr As rdoError
'Public RdoRes As adodb.recordset
'
'' Global Project Varibles
'Global glblActive As Label
Global gstrFilePath As String
Global gstrUser As String
Global gstrPassword As String
Global sSql As String
Global clsAdoCon As ClassFusionADO
Public bSqlRows As Boolean
Public sProcName As String

'
'Public gblnSqlRows As Boolean
Public sInitials As String       ' Esi2000 login
'Global gblnTime As Boolean ' Enable display time charges
'Public bSqlRows As Boolean
Global gblnUserAction As Boolean
Public bSysCalendar As Byte
'
'' Menu Constants
'Public Const MF_BYPOSITION = &H400&
'Public Const MF_GRAYED = &H1&
'Public Const SC_CLOSE = &HF060
'Public Const SC_MAXIMIZE = &HF030
'Public Const SC_MINIMIZE = &HF020
'Public Const SC_MOVE = &HF010
'Public Const SC_RESTORE = &HF120
'Public bRightArrowAsTab As Byte
Public Const SO_NUM_FORMAT = "000000"
'
'
'' Form Constants
'Public Const ES_RESIZE = 0
'Public Const ES_DONTRESIZE = -1
'Public Const ES_LIST = 0
'Public Const ES_DONTLIST = -1
'Public Const ES_IGNOREDASHES = 1 'Compress routine
'
'' MsgBox
'Public Const ES_NOQUESTION = &H124 'Question and return (Default NO)
'Public Const ES_YESQUESTION = &H24 'Question and return (Default YES)
'
'' StrCase funtion contstants
'Public Const ES_FIRSTWORD As Byte = 1
'
' Cursor types
Public Const ES_FORWARD = 0 'Default
Public Const ES_KEYSET = 1
Public Const ES_DYNAMIC = 2
Public Const ES_STATIC = 3
'
'' SetWindowPos Support
'Public Const Swp_NOMOVE = 2
'Public Const Swp_NOSIZE = 1
'Public Const Flags = Swp_NOMOVE Or Swp_NOSIZE
'Public Const hWnd_TopMost = -1
'Public Const Hwnd_NOTOPMOST = -2
'
'' Colors
'Global Const YELLOW = &HFFFF&
'
'' Grid
'Global Const ROWSPERPAGE = 8
'
'Global gbytScreen As Byte ' What screen we are on
'Global Const LOGIN = &H1
'Global Const SHOPS = &H2
'Global Const WCS = &H3
'Global Const jobs = &H4
'Global Const complete = &H5
'Global Const PKLIST = &H6
'Global Const lots = &H7
'
'
Global gstrEDIInFlPath As String
Global gstrEDIOutFlPath As String
Global gstrASNFName As String
Global gstrINVFName As String
Global gstrEDIArc As String
Global gstrProEDIArc As String

Global gstrPROEDIImpFP As String
Global gstrPROEDIExpFP As String

Global gstrRawEDIDir As String
Global gstrRawEDIFile As String
Global gstrRawEDIArch As String

Public ES_SYSDATE As Variant 'Server Date/Time to reduce calls

Public bDataHasChanged As Boolean
'Public sProcName As String
'Public Es_frmKeyDown(20) As New EsiKeyBd
'Public sRegistryAppTitle As String
Public bInsertOn As Boolean
Public bUserAction As Boolean
Public bEnterAsTab As Byte
'Public iBarOnTop As Byte
'
'Declare Function GetPrivateProfileString Lib "kernel32" Alias _
'"GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
'ByVal lpKeyName As Any, ByVal lpDefault As String, _
'ByVal lpReturnedString As String, ByVal nSize As Long, _
'ByVal lpFileName As String) As Long
'
Public Sub Main()
   Dim intResponse As Integer
   Dim sMsg As String
   Dim bRegOK As Boolean

   Dim strSaAdmin As String
   Dim strPW As String
   Dim strServer As String
   Dim strDataBase As String


   '*
   '* This will need to change (nth)
   '* gstrUser and gstrPassword
   '*

   gstrUser = "EdiUser"
   gstrPassword = "EdiUser"
   
   sProgName = "Fusion EDI"

   'get Esi2000 user who last logged in on this machine
   sInitials = Trim(GetSetting("Esi2000", "System", "UserInitials", sInitials))

   '    gstrServer = UCase$( _
   '        GetSetting("Esi2000", _
   '        "System", "ServerId", _
   '        gstrServer))

   gstrFilePath = GetSetting("Esi2000", _
                  "System", "FilePath", _
                  gstrFilePath)


   Dim strIniPath As String
   strIniPath = App.Path & "\" & "EDIInit.ini" '"d:\temp\VBA_MOSS.ini"

   InitGlobalIniInfo strIniPath

   Dim iTimeOut As Integer
'   strServer = GetSectionEntry("EDI_DEFAULTS", "DBSERVER_NAME", strIniPath)
'   strSaAdmin = GetSectionEntry("EDI_DEFAULTS", "DB_USERNAME", strIniPath)
'   strPW = GetSectionEntry("EDI_DEFAULTS", "DB_PASSWORD", strIniPath)
'   strDataBase = GetSectionEntry("EDI_DEFAULTS", "DB_NAME", strIniPath)

   ' Initialize the global values
'   sSaAdmin = strSaAdmin
'   sSaPassword = strPW
'   sserver = strServer
'   sDataBase = strDataBase
'
   sSysCaption = "EDI"

'   Set RdoEnv = rdoEnvironments(0)
'   RdoEnv.CursorDriver = rdUseIfNeeded
'
'   Set RdoCon = New rdoConnection
'   clsAdoCon.QueryTimeout = 60
'   iTimeOut = SQLSetConnectOption(clsAdoCon.hdbc, SQL_PRESERVE_CURSORS, SQL_PC_ON)
'   clsAdoCon.Connect = "UID=" & strSaAdmin & ";PWD=" & strPW & ";DRIVER={SQL Server};" _
'                    & "SERVER=" & strServer & ";DATABASE=" & strDataBase & ";"
'   clsAdoCon.EstablishConnection rdDriverNoPrompt

   'clsAdoCon.ExecuteSql "alter table Version add TestDatabase tinyint not null default 0"
   'Err.Clear
'   iTimeOut = clsAdoCon.QueryTimeout
'   If iTimeOut < 60 Then clsAdoCon.QueryTimeout = 60

'   gstrEDIInFlPath = GetSectionEntry("EDI_DEFAULTS", "EDI_INPUT_FILEPATH", strIniPath)
'   gstrEDIOutFlPath = GetSectionEntry("EDI_DEFAULTS", "EDI_OUTPUT_FILEPATH", strIniPath)
'   gstrASNFName = GetSectionEntry("EDI_DEFAULTS", "EDI_ASN_FNAME", strIniPath)
'   gstrINVFName = GetSectionEntry("EDI_DEFAULTS", "EDI_INV_FNAME", strIniPath)
'   gstrEDIArc = GetSectionEntry("EDI_DEFAULTS", "EDI_ARCHIVE_FILEPATH", strIniPath)
'   gstrProEDIArc = GetSectionEntry("EDI_DEFAULTS", "EDI_PROARCHIVE_FILEPATH", strIniPath)
'   gstrPROEDIImpFP = GetSectionEntry("EDI_DEFAULTS", "PROEDI_IMPORT_FP", strIniPath)
'   gstrPROEDIExpFP = GetSectionEntry("EDI_DEFAULTS", "PROEDI_EXPORT_FP", strIniPath)

'   If (Command = "IMPORT_EDI") Then
'      Dim strFileName As String
'      strFileName = Dir(gstrEDIInFlPath & "\*.*", vbDirectory)   ' Retrieve the first entry.
'      Do While strFileName <> ""
'         ' Ignore the current directory and the encompassing directory.
'         If strFileName <> "." And strFileName <> ".." Then
'            ' Import all the files from folders
'            ImpEDISalesOrder (gstrEDIInFlPath & "\"), strFileName
'            Debug.Print strFileName   ' Display entry only if it
'         End If
'         strFileName = Dir   ' Get next entry.
'      Loop
   On Error GoTo whoops    '@@@

   Dim strFileName As String
   Dim strSrcFPath As String
   Dim strDesFPath As String
   Dim strFileStamped As String
   Dim strDesArcFp As String
   Dim strProDesFPath As String
   Dim strFileEx As String
   Dim strTimeStamp As String
   Dim strInDate As String
   
   'Set clsAdoCon = New ClassFusionADO
   OpenDBServer
   
   Dim arv() As String
   If (Command <> "") Then
      arv = Split(Command, ",")
      Dim iLen As Integer
      iLen = UBound(arv) - LBound(arv) + 1

      If (iLen > 1) Then
         strInDate = Trim(arv(1))
      Else
         strInDate = ""
      End If

      'If (Command = "IMPORT_EDI") Then
      If (arv(0) = "IMPORT_EDI") Then
         Err.Clear

MsgBox "@@@ import from dir gstrPROEDIImpFP " & gstrPROEDIImpFP
         strFileName = Dir(gstrPROEDIImpFP & "\*.*", vbDirectory)   ' Retrieve the first entry.
         Do While strFileName <> ""
            ' Ignore the current directory and the encompassing directory.
            If strFileName <> "." And strFileName <> ".." Then
               ' Copy the file to the archive
               FindFileExten strFileName, strFileEx
'               strTimeStamp = Format(Now(), "mm-dd-yy.h.mm.ssA/P")
'               strFileStamped = strFileName & "_" & strTimeStamp & "_" & strFileEx
               strTimeStamp = Format(Now(), "mm-dd-yy.h.mm.ssA/P")
               strTimeStamp = Replace(strTimeStamp, "-", "")
               strTimeStamp = Replace(strTimeStamp, ".", "")
               
               'strFileStamped = strFileName & "(" & strTimeStamp & ")" & strFileEx
               strFileStamped = Replace(strFileName, ".EDI", "") & strTimeStamp & strFileEx

               ' Copy the file to the Fusion server
               strSrcFPath = gstrPROEDIImpFP & "\" & strFileName
               strDesFPath = gstrEDIInFlPath & "\" & strFileStamped
MsgBox "@@@ import copy strSrcFPath to strDesFPath " & strSrcFPath & " " & strDesFPath
               FileCopy strSrcFPath, strDesFPath


               strDesArcFp = gstrEDIArc & "\" & strFileStamped
MsgBox "@@@ import copy2 strSrcFPath to strDesArcFp " & strSrcFPath & " " & strDesArcFp
               FileCopy strSrcFPath, strDesArcFp

               ' If no error remove the file
               If (Err.Number = 0) Then
                  strSrcFPath = gstrPROEDIImpFP & "\" & strFileName
                  strProDesFPath = gstrProEDIArc & "\" & strFileStamped
                  ' Copy the file
MsgBox "@@@ import copy3 strSrcFPath to strProDesFPath " & strSrcFPath & " " & strProDesFPath
                  FileCopy strSrcFPath, strProDesFPath
                  ' Delete the source file
MsgBox "@@@ import delete strSrcFPath " & strSrcFPath
                  Kill strSrcFPath
               End If
            End If

            strFileName = Dir   ' Get next entry.
         Loop
      ElseIf (arv(0) = "CREATE_ASNOUT") Then
         ' Create ASN output
         CreateASNOut (gstrEDIOutFlPath & "\"), gstrASNFName, strInDate
      ElseIf (arv(0) = "CREATE_INVOUT") Then
         ' Create INV output file
         CreateInvoiceEDIFile (gstrEDIOutFlPath & "\"), gstrINVFName, strInDate
      ElseIf (arv(0) = "EXPORT_EDI") Then
         ' Create INV output file
'MsgBox "@@@ CreateInvoiceEDIFile " & gstrEDIOutFlPath
         CreateInvoiceEDIFile (gstrEDIOutFlPath & "\"), gstrINVFName, strInDate
         ' Create ASN output
'MsgBox "@@@ CreateASNOut " & gstrEDIOutFlPath
         CreateASNOut (gstrEDIOutFlPath & "\"), gstrASNFName, strInDate

         Err.Clear
         strFileName = Dir(gstrEDIOutFlPath & "\*.*", vbDirectory)   ' Retrieve the first entry.
         Do While strFileName <> ""
            ' Ignore the current directory and the encompassing directory.
            If strFileName <> "." And strFileName <> ".." Then
'MsgBox "@@@ strfilename=" & strFileName
               ' Copy the file to the Fusion server
               strSrcFPath = gstrEDIOutFlPath & "\" & strFileName
               strDesFPath = gstrPROEDIExpFP & "\" & strFileName
'MsgBox "@@@ filecopy " & strSrcFPath & " to " & strDesFPath
               FileCopy strSrcFPath, strDesFPath

               ' Copy the file to the archive
               FindFileExten strFileName, strFileEx
               strTimeStamp = Format(Now(), "mm-dd-yy.h.mm.ssA/P")

               strFileStamped = strFileName & "_" & strTimeStamp & "_" & strFileEx
               strDesArcFp = gstrEDIArc & "\" & strFileStamped
'MsgBox "@@@ filecopy " & strSrcFPath & " to " & strDesArcFp
               FileCopy strSrcFPath, strDesArcFp

               ' If no error remove the file
               If (Err.Number = 0) Then
                  ' Delete the source file
'MsgBox "@@@ delete " & strSrcFPath
                  Kill strSrcFPath
               End If
            End If

            strFileName = Dir   ' Get next entry.
         Loop

      ElseIf (arv(0) = "APPEND_EDIFILES") Then
         Err.Clear
         Dim strOutFile As String
         Dim nOutFNum As Integer

         strTimeStamp = Format(Now(), "mm-dd-yy.h.mm.ssA/P")
         strOutFile = gstrRawEDIDir & "\" & gstrRawEDIFile     ' C:\PROEDIW\EDIIN\UA0000.000
         nOutFNum = FreeFile
         
MsgBox "@@@ append open strOutFile " & strOutFile

         Open strOutFile For Output As nOutFNum

         If EOF(nOutFNum) Then
            strFileName = Dir(gstrRawEDIDir & "\*.*", vbDirectory)   ' Retrieve the first entry.

            Do While strFileName <> ""
               ' Ignore the current directory and the encompassing directory.
               If strFileName <> "." And strFileName <> ".." _
                  And strFileName <> gstrRawEDIFile Then
                  ' Copy the file to the Fusion server

                  Dim strInputFP As String
                  Dim nInputFNum As Integer
                  Dim sText As String
                  Dim sNextLine As String
                  Dim lLineCount As Long

                  strInputFP = gstrRawEDIDir & "\" & strFileName  'C:\PROEDIW\EDIIN\...
                  ' Get a free file number
                  nInputFNum = FreeFile
MsgBox "@@@ append open strInputFP " & strInputFP
                  Open strInputFP For Input As nInputFNum
                  ' Read the contents of the file
                  Do While Not EOF(nInputFNum)
                     Line Input #nInputFNum, sNextLine

                     If EOF(nOutFNum) Then
                        Print #nOutFNum, sNextLine
                        Debug.Print sNextLine
                     End If

                  Loop
                  Close nInputFNum

                  strFileStamped = strFileName & "_" & strTimeStamp & "_"
                  strDesArcFp = gstrRawEDIArch & "\" & strFileStamped
MsgBox "@@@ append copy strInputFP to strDesArcFp " & strInputFP & " " & strDesArcFp
                  FileCopy strInputFP, strDesArcFp 'copy to C:\EDI-ARCHIVE\Inbound

               End If
               strFileName = Dir   ' Get next entry.
            Loop

            ' Close the ouput file
            Close nOutFNum

            strFileStamped = gstrRawEDIFile & "(" & strTimeStamp & ").000"
            strDesArcFp = gstrRawEDIArch & "\" & strFileStamped
MsgBox "@@@ append copy2 strOutFile to strDesArcFp " & strOutFile & " " & strDesArcFp
            FileCopy strOutFile, strDesArcFp

            strDesArcFp = gstrEDIArc & "\" & strFileStamped

MsgBox "@@@ append copy3 strOutFile to strDesArcFp " & strOutFile & " " & strDesArcFp
            FileCopy strOutFile, strDesArcFp

         End If
      Else
         EDIMain.Show
      End If
      'MsgBox "end"
      End
   Else
         EDIMain.Show
   End If
   Exit Sub

whoops:
MsgBox "@@@ Error " & CStr(Err.Number) & " " & Err.Description

   ' MM OpenSqlServer

End Sub

Sub MouseCursor(MCursor As Integer)
   'Allows consistant MousePointer Updates
   Screen.MousePointer = MCursor
   gblnUserAction = True
End Sub

'Public Sub UpdateTables()
'
'End Sub
'
'
'Public Function GetDataSet(ssql,RdoDataSet As adodb.recordset, Optional iCursorType As Integer) As Boolean
'   ' Use local error Trapping
'   clsadocon.QueryTimeout = 40
'
'   If iCursorType = ES_FORWARD Then
'      'Forward only "cursor" (not a cursor)
'      Set RdoDataSet = clsadocon.OpenResultset(sSql, rdOpenForwardOnly, rdConcurReadOnly)
'      If Not RdoDataSet.BOF And Not RdoDataSet.EOF Then
'         GetDataSet = True
'      Else
'         GetDataSet = False
'      End If
'   Else
'      If iCursorType = ES_KEYSET Then
'         'Keyset cursor for Editing
'         Set RdoDataSet = clsadocon.OpenResultset(sSql, rdOpenKeyset, rdConcurRowVer)
'         If Not RdoDataSet.BOF And Not RdoDataSet.EOF Then
'            GetDataSet = True
'         Else
'            GetDataSet = False
'         End If
'      ElseIf iCursorType = ES_DYNAMIC Then
'         'Dynamic
'         Set RdoDataSet = clsadocon.OpenResultset(sSql, rdOpenDynamic, rdConcurRowVer)
'         If Not RdoDataSet.BOF And Not RdoDataSet.EOF Then
'            GetDataSet = True
'         Else
'            GetDataSet = False
'         End If
'      ElseIf iCursorType = ES_STATIC Then
'         'Static Cursor - Note: Needed for BLOBS
'         Set RdoDataSet = clsadocon.OpenResultset(sSql, rdOpenStatic, rdConcurReadOnly)
'         If Not RdoDataSet.BOF And Not RdoDataSet.EOF Then
'            GetDataSet = True
'         Else
'            GetDataSet = False
'         End If
'      End If
'      If Err > 0 Then GetDataSet = False
'   End If
'
'End Function
'
''Standard procedure for receiving resultsets
''bSqlRows = GetQuerySet (RdoRes, RdoQry, ES_FORWARD)
''See also GetDataSet for general queries
'
'Public Function GetQuerySet(RdoDataSet As adodb.recordset, RdoQueryDef As rdoQuery, Optional iCursorType As Integer) As Boolean
'   ' Use local error Trapping
'   clsadocon.QueryTimeout = 40
'   If iCursorType = ES_FORWARD Then
'      'Forward only "cursor" (not a real cursor)
'      Set RdoDataSet = RdoQueryDef.OpenResultset(rdOpenForwardOnly, rdConcurReadOnly)
'      If Not RdoDataSet.BOF And Not RdoDataSet.EOF Then
'         GetQuerySet = 1
'      Else
'         GetQuerySet = 0
'      End If
'   Else
'      'Keyset cursor for Editing
'      If iCursorType = ES_KEYSET Then
'         Set RdoDataSet = RdoQueryDef.OpenResultset(rdOpenKeyset, rdConcurRowVer)
'         If Not RdoDataSet.BOF And Not RdoDataSet.EOF Then
'            GetQuerySet = 1
'         Else
'            GetQuerySet = 0
'         End If
'      Else
'         'Static Cursor - Note: Needed for BLOBS
'         Set RdoDataSet = RdoQueryDef.OpenResultset(rdOpenStatic, rdConcurReadOnly)
'         If Not RdoDataSet.BOF And Not RdoDataSet.EOF Then
'            GetQuerySet = 1
'         Else
'            GetQuerySet = 0
'         End If
'      End If
'      If Err > 0 Then GetQuerySet = 0
'   End If
'End Function
'
Public Function Compress( _
                         TestNo As Variant, _
                         Optional iLength As Integer, _
                         Optional bIgnoreDashes As Byte) As String

   Dim A As Integer
   Dim K As Integer
   Dim PartNo As String
   Dim NewPart As String

   On Error GoTo modErr1
   PartNo = Trim$(TestNo)
   A = Len(PartNo)
   If A > 0 Then
      For K = 1 To A
         If bIgnoreDashes Then
            If Mid$(PartNo, K, 1) <> Chr$(32) And Mid$(PartNo, K, 1) <> Chr$(9) _
                    And Mid$(PartNo, K, 1) <> Chr$(39) Then
               NewPart = NewPart & Mid$(PartNo, K, 1)
            End If
         Else
            If Mid$(PartNo, K, 1) <> Chr$(45) And Mid$(PartNo, K, 1) <> Chr$(32) _
                    And Mid$(PartNo, K, 1) <> Chr$(9) And Mid$(PartNo, K, 1) <> Chr$(39) Then
               NewPart = NewPart & Mid$(PartNo, K, 1)
            End If
         End If
      Next
   End If
   If iLength > 0 Then
      If Len(NewPart) > iLength Then
         Beep
         NewPart = Left$(NewPart, iLength)
      End If
   End If
   Compress = NewPart
   Exit Function

modErr1:
   Resume modErr2
modErr2:
   On Error Resume Next
   Compress = TestNo
End Function

'
'Public Sub DoModuleErrors(frm As Form)
'   Dim RdoEvent As adodb.recordset
'   Dim bByte As Byte
'   Dim bInstr As Byte
'   '   Dim iFreeFile    As Integer
'   Dim iWarningType As Integer
'   Dim sMsg As String
'   Dim sMsg2 As String
'
'   'error log
'   Dim sDate As String * 16
'   Dim sSection As String * 8
'   Dim sForm As String * 12
'   Dim sErrNum As String * 10
'   Dim sErrSev As String * 2
'   Dim sProc As String * 10
'   Dim sUserName As String * 20
'   Dim ErringSQL As String
'
'   ErringSQL = sSql        'save sql - it gets overwritten below
'   MouseCursor 13
'   LockWindowUpdate 0
'   On Error Resume Next
'   sDate = Format(GetServerDateTime, "mm/dd/yy hh:mm")
'   sSection = Left(sProgName, 5)
'   sForm = Trim(frm.Name)
'   sErrNum = str(CurrError.Number)
'   sUserName = Left(cUR.CurrentUser, 18)
'
'   'Default Warning Flag. Setting bByte to True changes
'   'the Warning Flag and smooth closes the app if req'd.
'   iWarningType = vbExclamation
'   Select Case CurrError.Number
'      Case 3
'         sMsg = "Return Without GoSub"
'         bByte = 0
'      Case 5
'         sMsg = "Invalid Procedure Call"
'         bByte = 0
'      Case 6
'         sMsg = "Overflow"
'         bByte = 0
'      Case 7
'         sMsg = "Out Of Memory"
'         bByte = 1
'      Case 9
'         sMsg = "Subscript Out Of Range"
'         bByte = 0
'      Case 10
'         sMsg = "This Array Is Fixed Or Temporarily Locked"
'         bByte = 1
'      Case 11
'         sMsg = "Division By Zero"
'         bByte = 1
'      Case 13
'         sMsg = "Type Mismatch"
'         bByte = 0
'      Case 14
'         sMsg = "Out Of String Space"
'         bByte = 0
'      Case 16
'         sMsg = "Expression Too Complex"
'         bByte = 0
'      Case 17
'         sMsg = "Can't Perform Requested Operation"
'         bByte = 0
'      Case 18
'         sMsg = "User Interrupt Occurred"
'         bByte = 1
'      Case 20
'         sMsg = "Resume Without Error"
'         bByte = 0
'      Case 28
'         sMsg = "Out Of Strack Space"
'         bByte = 1
'      Case 35
'         sMsg = "Sub, Function, Or Property Not Defined"
'         bByte = 0
'      Case 47
'         sMsg = "Too Many DLL Application Clients"
'         bByte = 1
'      Case 48
'         sMsg = "Error In Loading DLL"
'         bByte = 1
'      Case 49
'         sMsg = "Bad DLL Calling Convention"
'         bByte = 1
'      Case 51
'         sMsg = "Internal Error"
'         bByte = 1
'      Case 52
'         sMsg = "Bad File Name Or Number"
'         bByte = 1
'      Case 53
'         sMsg = "File Not Found"
'         bByte = 0
'         iWarningType = vbInformation
'      Case 54
'         sMsg = "Bad File Mode"
'         bByte = 0
'      Case 55
'         sMsg = "File Already Open"
'         bByte = 0
'      Case 57
'         sMsg = "Device I/O Error"
'         bByte = 0
'      Case 58
'         sMsg = "File Already Exists"
'         bByte = 0
'      Case 59
'         sMsg = "Bad Record Length"
'         bByte = 1
'      Case 61
'         sMsg = "Disk Full"
'         bByte = 1
'      Case 62
'         sMsg = "Input Past End Of File"
'         bByte = 0
'      Case 63
'         sMsg = "Bad Record Number"
'         bByte = 0
'      Case 67
'         sMsg = "Too Many Files"
'         bByte = 1
'      Case 68
'         sMsg = "Device Unavailable"
'         bByte = 0
'      Case 70
'         sMsg = "Permission Denied"
'         bByte = 0
'      Case 71
'         sMsg = "Disk Not Ready"
'         bByte = 0
'      Case 74
'         sMsg = "Can't Rename With Different Drive"
'         bByte = 0
'      Case 75
'         sMsg = "Path/File Access Error"
'         bByte = 0
'      Case 76
'         sMsg = "Path Not Found." & vbCrLf _
'                & "An Attempt Was Made To Open A File From An" & vbCrLf _
'                & "Invalid Location. Check The User Filepath Settings."
'         bByte = 0
'      Case 91
'         sMsg = "Object Variable Or With Block Variable Not Set."
'         sMsg = sMsg & vbCrLf & "Check Network Connection"
'         bByte = 0
'      Case 94
'         sMsg = "Invalid Use Of Null. Please Report This Error."
'         bByte = 0
'      Case 340
'         sMsg = "Control Element Does Not Exist."
'         bByte = 0
'      Case 380
'         sMsg = "Invalid Property Value"
'         bByte = 0
'      Case 438
'         sMsg = "Function Not Supported By Object.  Please Report This Error."
'         bByte = 0
'      Case 482, 483, 486
'         sMsg = "Printer Error"
'         bByte = 0
'         'Jet
'      Case 3001 To 3648
'         sMsg = "JET DSS Database Error. Contact Systems Administrator."
'         bByte = 0
'         'Crystal
'      Case 20500
'         sMsg = "Not Enough Memory To Complete Report. " & vbCrLf _
'                & "SSCSDK32.DLL May Be Missing Or Corrupt."
'         bByte = 0
'      Case 20501 To 20506
'         sMsg = "Undocumented Crystal Reports Error. " & vbCrLf _
'                & "Contact Systems Administrator."
'         CurrError.Description = Left(sMsg, 34)
'         bByte = 0
'      Case 20507
'         sMsg = "Report Wasn't Found Or Couldn't Be Loaded. " & vbCrLf _
'                & "Check Your Report Path In Settings."
'         CurrError.Description = "Couldn't Find The Requested Report."
'         bByte = 0
'      Case 20508, 20509, 20511, 20512, 20514
'         sMsg = "Undocumented Crystal Reports Error. " & vbCrLf _
'                & "Contact Systems Administrator."
'         CurrError.Description = Left(sMsg, 34)
'         bByte = 0
'      Case 20510
'         'Crystal does not broadcast the name
'         sMsg = "The Requested Formula Doesn't Exist (Invalid Formula Name). "
'         CurrError.Description = Left(sMsg, 35)
'         bByte = 0
'      Case 20513
'         sMsg = "The Requested Printer Is Not Valid. " & vbCrLf _
'                & "Contact Systems Administrator."
'         CurrError.Description = Left(sMsg, 34)
'         bByte = 0
'      Case 20515
'         bInstr = InStr(1, CurrError.Description, "<Record_Selection>")
'         If bInstr > 0 Then
'            CurrError.Description = "Error In The SQL Query (Record Selection)"
'            sMsg = "Error In SQL Statement  (Record Selection)" & vbCrLf _
'                   & "Contact Your Systems Administrator."
'         Else
'            bInstr = InStr(1, CurrError.Description, "<")
'            CurrError.Description = "Error In Formula " & Mid(CurrError.Description, bInstr)
'            sMsg = CurrError.Description & vbCrLf _
'                   & "Contact Your Systems Administrator."
'         End If
'         bByte = 0
'      Case 20516, 20517, 20519
'         sMsg = "Not Enough Windows Resources To Complete Report. " & vbCrLf _
'                & "Close Some Applications,  " & sSysCaption & " And Restart."
'         bByte = 0
'      Case 20518
'         sMsg = "An Attempt Was Made To Hide Or Show A Report Group" & vbCrLf _
'                & "That Does Not Exist. Contact Your Systems Administrator."
'         CurrError.Description = "Hide/Show A Group Does Not Exist"
'         bByte = 0
'      Case 20520
'         sMsg = "Print Job Started And Report In Progress Or " & vbCrLf _
'                & "There Is No Default Print Or The Printer Is Offline." & vbCrLf _
'                & "Crystal Reports Notice Not An Error."
'         iWarningType = vbInformation
'         bByte = 0
'      Case 20521 To 20522
'         sMsg = "Not Enough Windows Resources To Complete Report. " & vbCrLf _
'                & "Close Some Applications, " & sSysCaption & " And Restart."
'         bByte = 0
'      Case 20523
'         sMsg = "Invalid Character In The GroupConditioning." & vbCrLf _
'                & "Contact Your Systems Administrator."
'         CurrError.Description = Left(sMsg, 35)
'         bByte = 0
'      Case 20524
'         sMsg = "Undocumented Crystal Reports Error. " & vbCrLf _
'                & "Contact Systems Administrator."
'         CurrError.Description = Left(sMsg, 34)
'      Case 20525
'         sMsg = "Report Is Damaged. Unable To Open Report. " & vbCrLf _
'                & "Contact Your Systems Administrator."
'         bByte = 0
'      Case 20526
'         sMsg = "No Default Printer Has Been Set. "
'         CurrError.Description = sMsg
'         bByte = 0
'      Case 20527
'         sMsg = "Error In SQL Server Connection. " & vbCrLf _
'                & "The Report Cannot Resolve The Data."
'         CurrError.Description = "Crystal Reports Cannot Resolve The Data"
'         bByte = 0
'      Case 20529
'         sMsg = "Your Disk Drive Is Full And Files May Be Lost." & vbCrLf _
'                & "Exit " & sSysCaption & " And Free Resources."
'         bByte = 0
'      Case 20530
'         sMsg = "An Attempt Was Made To Access A Report In Use. " & vbCrLf _
'                & "Try The Report Again."
'         CurrError.Description = "An Attempt Was Made To Access A Report In Use"
'         iWarningType = vbInformation
'         bByte = 0
'      Case 20531
'         sMsg = "Incorrect Password. Permission Denied."
'         bByte = 0
'      Case 20532
'         sMsg = "File I/O Error. Disk Problem Other Than Full." & vbCrLf _
'                & "Exit " & sSysCaption & " Contact Systems Administrator."
'         bByte = 0
'      Case 20533
'         sMsg = "Unable To Open The Database File." & vbCrLf _
'                & "Contact Your Systems Administrator."
'         bByte = 0
'      Case 20534
'         sMsg = "The Database Columns Are Not Correctly Mapped" & vbCrLf _
'                & "Or Have A Collation Conflict. Contact Your" & vbCrLf _
'                & "Systems Administrator."
'         CurrError.Description = "Column Mapping Or Collation Problem."
'         bByte = 0
'      Case 20535 To 20543
'         sMsg = CurrError.Description
'         bByte = 0
'      Case 20544
'         sMsg = "This Report Is Open By Another User." & vbCrLf _
'                & "Try The Report Again In A Few Minutes."
'         bByte = 0
'      Case 20545
'         sMsg = CurrError.Description
'         iWarningType = vbInformation
'         bByte = 0
'      Case 20546 To 20598
'         sMsg = CurrError.Description
'         bByte = 0
'      Case 20599
'         sMsg = "ODBC Permissions/Access Error. Check ODBC Data Source." & vbCrLf _
'                & "DSN " & sDsn & " May Be Improperly Installed Or Does Not Exist."
'         bByte = 0
'      Case 20600 To 20996
'         sMsg = "Undocumented Crystal Reports Error." & vbCrLf _
'                & "Contact Your Systems Administrator."
'      Case 20997
'         sMsg = "Invalid Report Path Or No Network Permissions." & vbCrLf _
'                & "Check Your Report Path And Server Permissions."
'         bByte = 0
'      Case 20998
'         sMsg = "Report Path Is Too Long.     " & vbCrLf _
'                & "Use A Mapped Path Name Instead (x:\somedir) Or " _
'                & "Possible Mismatch Of Graph DLL Libraries."
'         bByte = 0
'         'Rdo
'      Case 40000
'         sMsg = "An Error Occurred Configuring The DataSource Name."
'         bByte = 0
'      Case 40001
'         sMsg = "SQL Returned No Data Found From Query."
'         iWarningType = vbInformation
'         bByte = 0
'      Case 40002
'         If Left(CurrError.Description, 5) = "22003" Then
'            sMsg = Left(CurrError.Description, 5) & "-Attempted Insert A Value Greater" & vbCrLf _
'                   & "Than The Requested Column Allows."
'            bByte = 0
'         Else
'            If Left(CurrError.Description, 5) = "S0002" Then
'               sMsg = Left(CurrError.Description, 5) & "-The Requested Table Wasn't Found."
'               bByte = 0
'            Else
'               If Left(CurrError.Description, 5) = "S0022" Then
'                  'sMsg = Left(CurrError.Description, 5) & "-The Requested Column Wasn't Found."
'                  sMsg = CurrError.Description
'                  bByte = 0
'               Else
'                  If Left(CurrError.Description, 5) = "01000" Then
'                     'sMsg = "An Attempt Was Made To Add A Duplicate Record" & vbCrLf _
'                     '       & "Or The Data To Be Inserted Is Not Valid."
'                     sMsg = CurrError.Description
'                     bByte = 0
'                  Else
'                     If InStr(CurrError.Description, "0851") > 0 Then
'                        sMsg = "ODBC Link Was Lost. Reconnection Required."
'                        bByte = 1
'                     ElseIf InStr(CurrError.Description, "S1T00") Then _
'                                     sMsg = "The Query Has Timed Out, Check Settings."
'                        bByte = 0
'                     ElseIf InStr(CurrError.Description, "S1010") Then _
'                                     sMsg = "The Query Has Timed Out. Bulk Copy May Be Turned On/"
'                        bByte = 0
'                     ElseIf InStr(CurrError.Description, "08S01") Then _
'                                     sMsg = "The Network Connection With SQL Server Failed."
'                        bByte = 0
'                     ElseIf InStr(CurrError.Description, "07S01") Then _
'                                     sMsg = "Invalid Default Parameter In Query Results."
'                        bByte = 0
'                        Err.Clear
'                        Exit Sub
'                     Else
'                        sMsg = "Internal ODBC Error Encountered."
'                        bByte = 0
'                     End If
'                  End If
'               End If
'            End If
'            If Left(CurrError.Description, 5) = "37000" Then
'               'Changed some for SSL 7.0
''               sMsg = Left(CurrError.Description, 5) & vbCrLf _
''                      & "The Cursor Is No Longer Open. Invalid Character " & vbCrLf _
''                      & "Found Or The Query Couldn't Process Requested Data. " & vbCrLf _
''                      & "Please Report This To Your Systems Administrator."
'               sMsg = Left(CurrError.Description, 5) & vbCrLf _
'                      & "SQL error. Invalid Character: " & vbCrLf _
'                      & CurrError.Description & vbCrLf _
'                      & "Please Report This To Your Systems Administrator."
'               bByte = 0
'            End If
'         End If
'      Case 40003
'         sMsg = "An Invalid Value For The Cursor Driver Was Passed."
'         bByte = 0
'      Case 40004
'         sMsg = "An Invalid ODBC Handle Was Encountered."
'         bByte = 0
'      Case 40005
'         sMsg = "Invalid Connection String."
'         bByte = 1
'      Case 40006
'         sMsg = "An Unexpected Error Occurred."
'         bByte = 0
'      Case 40008
'         sMsg = "Invalid Operation For Forward-Only Cursor."
'         bByte = 0
'      Case 40009
'         sMsg = "No Current Row (No Matching Query Data Found)."
'         iWarningType = vbInformation
'         bByte = 0
'      Case 40010
'         sMsg = "Invalid Row For Add New."
'         bByte = 0
'      Case 40011
'         sMsg = "Object Is Invalid Or Not Set."
'         bByte = 0
'      Case 40012
'         sMsg = "Invalid Seek Flag."
'         bByte = 0
'      Case 40013
'         sMsg = "Partial Equality Requires String Column."
'         bByte = 0
'      Case 40014
'         sMsg = "Incompatible Data Types For Compare."
'         bByte = 0
'      Case 40015
'         sMsg = "Can't Create Prepared Statement."
'         bByte = 0
'      Case 40016
'         sMsg = "Version.DLL Error."
'         bByte = 1
'      Case 40017, 40018
'         sMsg = "Can't Execute Statement."
'         bByte = 0
'      Case 40019
'         sMsg = "An Invalid Value For The Concurrency Option."
'         bByte = 0
'      Case 40020
'         sMsg = "Can't Open Result Set For Unnamed Table."
'         bByte = 0
'      Case 40021
'         sMsg = "Object Collection Error."
'         bByte = 0
'      Case 40022
'         sMsg = "The RDO Results Set Is Empty (No Data)."
'         iWarningType = vbInformation
'         bByte = 0
'      Case 40023
'         sMsg = "Invalid State For Cursor Move. "
'         bByte = 0
'      Case 40024
'         sMsg = "Already Beyond The End Of The Result Set."
'         bByte = 0
'      Case 40025
'         sMsg = "BOF Already Set."
'         bByte = 0
'      Case 40026
'         sMsg = "Invalid Result Set State For Update."
'         bByte = 0
'      Case 40027
'         sMsg = "Invalid Bookmark Or No Bookmark Allowed."
'         bByte = 0
'      Case 40028
'         sMsg = "Invalid Bookmark Argument To Move."
'         bByte = 0
'      Case 40029
'         sMsg = "Current Row As EOF/BOF Already Set."
'         bByte = 0
'      Case 40030
'         sMsg = "Already At BOF."
'         bByte = 0
'      Case 40031
'         sMsg = "Already At EOF."
'         bByte = 0
'      Case 40032
'         sMsg = "Couldn't Load The ODBC Installation Library."
'         bByte = 1
'      Case 40033
'         sMsg = "An Invalid Value For The Prompt Option Was Passed."
'         bByte = 1
'      Case 40034
'         sMsg = "An Invalid Value For The Cursor Type Parameter Was Passed."
'         bByte = 0
'      Case 40035
'         sMsg = "Column Not Bound Correctly."
'         bByte = 0
'      Case 40036
'         sMsg = "Unbound Column-Use Get Chunk Method."
'         bByte = 0
'      Case 40037
'         sMsg = "Can't Assign Value To Unbound Column."
'         bByte = 0
'      Case 40038
'         sMsg = "Can't Assign Value To Non-Updatable Field."
'         bByte = 0
'      Case 40039
'         sMsg = "Can't Assign Value To Column Unless In Edit Mode."
'         bByte = 0
'      Case 40040
'         sMsg = "Incorrect Type For Parameter."
'         bByte = 0
'      Case 40041
'         sMsg = "Object Collection: Couldn't Find Column Requested By Query."
'         bByte = 0
'      Case 40042
'         sMsg = "Can't Assign Value To Unbound Parameter."
'         bByte = 0
'      Case 40043
'         sMsg = "Can't Assign Value To Output-Only Parameter."
'         bByte = 0
'      Case 40044
'         sMsg = "Incorrect RDO Parameter Type."
'         bByte = 0
'      Case 40045
'         sMsg = "Tried To Execute A Query With An Asynchronous Query In Progress."
'         bByte = 0
'      Case 40046
'         sMsg = "The Object Has Already Been Closed."
'         bByte = 0
'      Case 40047
'         sMsg = "Invalid Name For The Environment."
'         bByte = 0
'      Case 40048
'         sMsg = "Environment Name Already Exists In The Collection."
'         bByte = 0
'      Case 40049
'         sMsg = "Object Collection Is Read-Only."
'         bByte = 0
'      Case 40050
'         sMsg = "Get New Enum: Couldn't Get Interface."
'         bByte = 0
'      Case 40051
'         sMsg = "Assignment To Count Property Not Allowed."
'         bByte = 0
'      Case 40052
'         sMsg = "You Must Use Append Chunk To Set Data In A Text Or Image."
'         bByte = 0
'      Case 40053
'         sMsg = "Object Collection: Can't Add Non Object Item."
'         bByte = 0
'      Case 40054
'         sMsg = "An Invalid Parameter Was Passed."
'         bByte = 0
'      Case 40055
'         sMsg = "Invalid Operation."
'         bByte = 0
'      Case 40056
'         sMsg = "The Row Has Been Deleted."
'         bByte = 0
'      Case 40057
'         sMsg = "An Attempt Was Made To Issue A Select Statement Using Execute."
'         bByte = 0
'      Case 40058
'         sMsg = "Can't Update Column, The Result Set Is Read Only."
'         bByte = 0
'      Case 40059
'         sMsg = "Cancel Has Been Selected In An ODBC Dialog Requesting Parameters."
'         iWarningType = vbInformation
'         bByte = 0
'      Case 40060
'         sMsg = "Needs Chunk Required Flags."
'         bByte = 0
'      Case 40061
'         sMsg = "Could Not Load Resource Library."
'         bByte = 1
'      Case 40069
'         sMsg = "General Client Cursor Error."
'         bByte = 0
'      Case 40071
'         sMsg = "The RDO Connection Object Is Not Connected To A Data Source."
'         bByte = 1
'      Case 40072
'         sMsg = "The RDO Connection Object Is Already Connected To The Data Source."
'         bByte = 0
'      Case 40073
'         sMsg = "The RDO Connection Object Is Busy Connecting " & vbCrLf _
'                & "To The Data Source. Retry The Selection."
'         bByte = 0
'      Case 40074
'         sMsg = "The RDO Query Or RDO Results Set Has No Active Connection Source."
'         bByte = 1
'      Case 40075
'         sMsg = "Incorrect Cursor Driver."
'         bByte = 0
'      Case 40076
'         sMsg = "This Property Is Currently Read Only."
'         iWarningType = vbInformation
'         bByte = 0
'      Case 40077
'         sMsg = "The Object Is Already In The Collection."
'         iWarningType = vbInformation
'         bByte = 0
'      Case 40078
'         sMsg = "Failed To Load RDOCURS.DLL"
'         bByte = 1
'      Case 40079
'         sMsg = "Can't Find The Requested Table To Update."
'         bByte = 0
'      Case 40080, 40081, 40082, 40083, 40085
'         sMsg = "Invalid RDO/SQL Server Option."
'         bByte = 0
'      Case 40088
'         sMsg = "No Open Cursor For Transaction Commit."
'         bByte = 0
'      Case 40500, 40501, 40502, 40503
'         sMsg = "Unexpected Internal RDO Error "
'         bByte = 1
'      Case 40504
'         sMsg = "Could Not Refresh Controls."
'         bByte = 0
'      Case 40505
'         sMsg = "Invalid Property Value."
'         bByte = 0
'      Case 40506
'         sMsg = "Invalid Collection Object."
'         bByte = 0
'      Case 40507
'         sMsg = "Method Cannot Be Called In RDO's Current State."
'         bByte = 0
'      Case 40508
'         sMsg = "One Or More Of The Arguments Is Invalid."
'         bByte = 0
'      Case 40509
'         sMsg = "Result Set Is Empty."
'         iWarningType = vbInformation
'         bByte = 0
'      Case 40510
'         sMsg = "Out Of Memory. Close " & sSysCaption & "."
'         bByte = 1
'      Case 40511
'         sMsg = "Result Set Not Available."
'         bByte = 0
'      Case 40512
'         sMsg = "The Connection Is Not Open."
'         bByte = 1
'      Case 40513, 40514
'         sMsg = "Property Cannot Be Set In RDC's Current State."
'         bByte = 0
'      Case 40515
'         sMsg = "Type Mismatch."
'         bByte = 0
'      Case 40516
'         sMsg = "Cannot Connect To Remote Data Object."
'         bByte = 1
'      Case Else
''         sMsg = "Undocumented Error           "
'         sMsg = CurrError.Description
'         bByte = 1
'   End Select
'
'   'add Crystal Reports filename, error description,
'   'formulas and SQL to message.
'   Select Case CurrError.Number
'         'if Sql Server error, add SQL to log
'      Case 40000 To 40516
'         sMsg2 = "SQL Server Command: " & ErringSQL
'
'      Case Else
'         sMsg2 = sSysCaption & " System."
'   End Select
'   If iWarningType = vbInformation Then
'      sMsg = "Notification" & str(CurrError.Number) & vbCrLf & sMsg & vbCrLf & sMsg2
'   Else
'      If bByte = 1 Then
'         sMsg = "Error" & str(CurrError.Number) & vbCrLf & sMsg & vbCrLf & sMsg2
'      Else
'         sMsg = "Warning" & str(CurrError.Number) & vbCrLf & sMsg & vbCrLf & sMsg2
'      End If
'   End If
'   MouseCursor 0
'   'Show the user and do as required
'   MDISect.Enabled = True
'
'   If bByte = 1 Then
'      sErrSev = "16"
'      '            If Len(Trim(CurrError.Description)) < 35 Then
'      '                Print #iFreeFile, sDate; sSection; sForm; sUserName; _
'      '                    sErrNum; sErrSev; " "; sProcName; " "; Trim(CurrError.Description)
'      '            Else
'      '                Print #iFreeFile, sDate; sSection; sForm; sUserName; _
'      '                    sErrNum; sErrSev; " "; sProcName; " "; Left$(CurrError.Description, 35)
'      '            End If
'      '            Close iFreeFile
'      sMsg = sMsg & vbCrLf & sProcName
'      sMsg = sMsg & vbCrLf & "Contact System Administrator"
'      MsgBox sMsg, vbCritical, frm.Caption
'   Else
'      If iWarningType = vbInformation Then sErrSev = "64" Else sErrSev = "48"
'      sMsg = sMsg & vbCrLf & "Procedure: " & sProcName
'      MsgBox sMsg, iWarningType, frm.Caption
'      sErrSev = "48"
'      CurrError.Number = 0
'   End If
'
'   Close
'   Err.Clear
'
'   'log the error in the SystemEvents table
'   'this must happen on a separate connection, otherwise, when the transaction is rolled back,
'   'the new SystemEvents row will be rolled back as well
'
'   sSql = "SELECT * FROM SystemEvents"
'   Dim rdoCon2 As rdoConnection
'   Set rdoCon2 = GetTemporarySqlConnection
'   Set RdoEvent = rdoCon2.OpenResultset(sSql, rdOpenKeyset, rdConcurRowVer)
'
'   With RdoEvent
'      .AddNew
'      !Event_Date = Format$(Now, "mm/dd/yy h:mm AM/PM")
'      !Event_Section = sSection
'      !Event_Form = sForm
'      !Event_User = sUserName
'      !Event_Event = Val(sErrNum)
'      !Event_Warning = Val(sErrSev)
'      !Event_Procedure = sProcName
'      !Event_Text = Left(Trim(sMsg), 4096)
'      .Update
'   End With
'   Set RdoEvent = Nothing
'   rdoCon2.Close
'   Set rdoCon2 = Nothing
'   If sErrSev = "16" Then
'      CloseFiles
'   End
'Else
'   sProcName = ""
'End If
'
'End Sub


Sub CloseFiles()
   On Error Resume Next
   Close
   'clsAdoCon.Close
   InvalidateRect 0&, 0&, False
   Set MDISect = Nothing
   End

End Sub

Public Sub FindCustomer(frm As Form, sCustomerNickname, Optional bNeedsMore As Byte)
   Dim CusRes As ADODB.Recordset
   On Error GoTo modErr1
   sSql = "Qry_FindCustomer '" & Compress(sCustomerNickname) & "' "
   bSqlRows = clsAdoCon.GetDataSet(sSql, CusRes)
   If bSqlRows Then
      With CusRes
         On Error Resume Next
         frm.lblCst = "" & Trim(!CUNICKNAME)
         frm.cmbCst = "" & Trim(!CUNICKNAME)
         frm.lblNme = "" & Trim(!CUNAME)
         frm.txtNme = "" & Trim(!CUNAME)
         If bNeedsMore Then
            frm.txtDis = Format(!CUDISCOUNT, "#0.00")
            frm.txtFra = Format(!CUFRTALLOW, ES_QuantityDataFormat)
            frm.txtFrd = Format(!CUFRTDAYS, "##0")
         End If
         ClearResultSet CusRes
      End With
   Else
      On Error Resume Next
      frm.lblNme = ""
      frm.txtNme = "*** Customer Wasn't Found ***"
      If Trim(frm.cmbCst) = "" Then frm.txtNme = "*** No Customer Selected ***"
   End If
   Set CusRes = Nothing
   Exit Sub

modErr1:
   sProcName = "findcustomer"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors frm

End Sub

Public Sub LoadComboBox(Cntrl As Control, Optional ColumnNumber As Integer)
   Dim ComboLoad As ADODB.Recordset
   Cntrl.Clear
   If sSql = "" Then Exit Sub
   ColumnNumber = ColumnNumber + 1
   'Set ComboLoad = clsAdoCon.OpenResultset(sSql, rdOpenForwardOnly, rdConcurReadOnly)
   Set ComboLoad = clsAdoCon.GetRecordSet(sSql, ES_STATIC)
   If Not ComboLoad.BOF And Not ComboLoad.EOF Then
      With ComboLoad
         Do Until .EOF
            'AddComboStr Cntrl.hwnd, "" & Trim(.rdoColumns(ColumnNumber))
            AddComboStr Cntrl.hwnd, "" & Trim(.Fields(ColumnNumber))
            .MoveNext
         Loop
         ClearResultSet ComboLoad
      End With
   End If
   If Cntrl.ListCount <> 0 Then
      bSqlRows = 1
      Cntrl.ListIndex = 0
   Else
      bSqlRows = 0
   End If
   sSql = ""
   Set ComboLoad = Nothing

End Sub

Public Sub AddComboStr(lhWnd As Long, sString As String)
   SendMessageStr lhWnd, CB_ADDSTRING, 0&, ByVal "" & Trim(sString)

End Sub

Sub FormLoad(frm As Form, Optional DontList As Boolean, Optional noResize As Boolean)
End Sub

'Public Sub LoadComboBoxAndSelect(cbo As ComboBox, Optional SelectString As String)
'
'   ' load the combobox with the contents of the first column returned from ssql
'   ' select the first entry >= the string specified
'   ' if no string is specified, select the first entry
'
'
'   Dim rdo As adodb.Recordset
'   cbo.Clear
'
'   'text is read-only for dropdown list
'   If cbo.Style <> 2 Then
'      cbo.Text = ""
'   End If
'   If sSql = "" Then Exit Sub
'   'Set rdo = clsAdoCon.OpenResultset(sSql, rdOpenForwardOnly, rdConcurReadOnly)
'   Set rdo = clsAdoCon.GetRecordSet(sSql, ES_FORWARD)
'   If Not rdo.BOF And Not rdo.EOF Then
'      bSqlRows = True
'      With rdo
'         Do Until .EOF
'            If cbo.ListCount = 32766 Then
'               AddComboStr cbo.hwnd, "MORE THAN 32767 ROWS"
'               'cbo.ListIndex = cbo.ListCount - 1
'               cbo.Text = "MORE THAN 32767 ROWS"
'               Exit Do
'            End If
'            AddComboStr cbo.hwnd, "" & Trim(.rdoColumns(0))
'            If cbo.Text = "" Then
'               If Trim(.rdoColumns(0)) >= SelectString Then
'                  If .rdoColumns(0) <> "<ALL>" Then
'                     'cbo.Text = Trim(.rdoColumns(0))
'                     cbo.ListIndex = cbo.ListCount - 1
'                  End If
'               End If
'            End If
'            .MoveNext
'         Loop
'         ClearResultSet rdo
'      End With
'   Else
'      bSqlRows = False
'   End If
'   Set rdo = Nothing
'
'   'If cbo.ListIndex = -1 And cbo.ListCount > 0 Then
'   '   cbo.ListIndex = 0      'CAUSES COLLAPSE
'   'End If
'
'
'End Sub
'
Public Sub LoadComboBoxAndSelect(cbo As ComboBox, Optional SelectString As String)

   ' load the combobox with the contents of the first column returned from ssql
   ' select the first entry >= the string specified
   ' if no string is specified, select the first entry
   
      
   Dim Ado As ADODB.Recordset
   cbo.Clear
   
   'text is read-only for dropdown list
   If cbo.Style <> 2 Then
      cbo.Text = ""
   End If
   If sSql = "" Then Exit Sub
   Set Ado = clsAdoCon.GetRecordSet(sSql, ES_FORWARD)
 
'   Set rdo = clsadocon.OpenResultset(sSql, rdOpenForwardOnly, rdConcurReadOnly)
   If Not Ado.BOF And Not Ado.EOF Then
      bSqlRows = True
      With Ado
         Do Until .EOF
            If cbo.ListCount = 32766 Then
               AddComboStr cbo.hwnd, "MORE THAN 32767 ROWS"
               'cbo.ListIndex = cbo.ListCount - 1
               cbo.Text = "MORE THAN 32767 ROWS"
               Exit Do
            End If
            AddComboStr cbo.hwnd, "" & Trim(.Fields(0))
            If cbo.Text = "" Then
               If Trim(.Fields(0)) >= SelectString Then
                  If .Fields(0) <> "<ALL>" Then
                     'cbo.Text = Trim(.Fields(0))
                     cbo.ListIndex = cbo.ListCount - 1
                  End If
               End If
            End If
            .MoveNext
         Loop
         ClearResultSet Ado
      End With
   Else
      bSqlRows = False
   End If
   Set Ado = Nothing
   
   'If cbo.ListIndex = -1 And cbo.ListCount > 0 Then
   '   cbo.ListIndex = 0      'CAUSES COLLAPSE
   'End If
   
   
End Sub




Public Sub ShowCalendar(frm As Form, Optional iAdjust As Integer)

   'display date selection calendar

   Dim iAdder As Integer
   Dim lLeft As Long
   Dim lTop As Long
   Dim sDate As Date

   Dim combo As Control
   Set combo = frm.ActiveControl


   If IsDate(frm.ActiveControl.Text) Then
      'frm.ActiveControl.AddItem frm.ActiveControl.Text
   Else
      'frm.ActiveControl.AddItem Format(Now, "mm/dd/yy")
      combo.Text = Format(Now, "mm/dd/yy")
   End If

   'set form to pass date back to
   Set SysCalendar.FromForm = frm

   'On Error Resume Next
   'See if there is a date in the combo
   If IsDate(frm.ActiveControl.Text) Then
      sDate = Format(frm.ActiveControl.Text, "mm/dd/yy")
   Else
      sDate = Format(ES_SYSDATE, "mm/dd/yy")
   End If


   ' MM SysCalendar.Move lLeft, lTop

   bSysCalendar = True
   If IsDate(sDate) Then SysCalendar.Calendar1.Value = Format(sDate, "mm/dd/yyyy")

   'if parent form is modal, we must show this as modal too
   On Error Resume Next
   SysCalendar.Calendar1.Refresh
   DoEvents
   SysCalendar.Show
   If Err Then
      SysCalendar.Show vbModal
   End If
   'refresh it so that it doesn't blink out
   SysCalendar.Calendar1.Refresh
   'combo.Refresh

End Sub

Public Function CheckLen(sTextBox As String, iTextLength As Integer) As String
   sTextBox = Trim(sTextBox)
   If Len(sTextBox) > iTextLength Then sTextBox = Left(sTextBox, iTextLength)
   CheckLen = sTextBox
   iTextLength = InStr(1, CheckLen, Chr$(39))
   If iTextLength > 0 Then CheckLen = ReplaceString(CheckLen)

End Function

Public Function ReplaceString(ByVal OldString As String) As String
   Dim NewString As String
   'Quotation with alternate
   NewString = Replace(OldString, Chr$(34), Chr$(146) & Chr$(146))
   'Apostrophe with alternate
   NewString = Replace(NewString, Chr$(39), Chr$(146))
   ReplaceString = NewString

End Function

'
'
Public Function GetSectionEntry(ByVal strSectionName As String, ByVal strEntry As String, ByVal strIniPath As String) As String


   Dim X As Long
   Dim sSection As String, sEntry As String, sDefault As String
   Dim sRetBuf As String, iLenBuf As Integer, sFileName As String
   Dim sValue As String

   On Error GoTo modErr1

   sSection = strSectionName
   sEntry = strEntry
   sDefault = ""
   sRetBuf = String(256, vbNull) '256 null characters
   iLenBuf = Len(sRetBuf)
   sFileName = strIniPath
   X = GetPrivateProfileString(sSection, sEntry, _
                     "", sRetBuf, iLenBuf, sFileName)
   sValue = Trim(Left$(sRetBuf, X))

   If sValue <> "" Then
      GetSectionEntry = sValue
   Else
      GetSectionEntry = ""
   End If

   Exit Function

modErr1:
   sProcName = "GetSectionEntry"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description

End Function


Public Function InitGlobalIniInfo(strIniPath As String)

   gstrEDIInFlPath = GetSectionEntry("EDI_DEFAULTS", "EDI_INPUT_FILEPATH", strIniPath)
   gstrEDIOutFlPath = GetSectionEntry("EDI_DEFAULTS", "EDI_OUTPUT_FILEPATH", strIniPath)
   gstrASNFName = GetSectionEntry("EDI_DEFAULTS", "EDI_ASN_FNAME", strIniPath)
   gstrINVFName = GetSectionEntry("EDI_DEFAULTS", "EDI_INV_FNAME", strIniPath)
   gstrProEDIArc = GetSectionEntry("EDI_DEFAULTS", "EDI_PROARCHIVE_FILEPATH", strIniPath)
   gstrEDIArc = GetSectionEntry("EDI_DEFAULTS", "EDI_ARCHIVE_FILEPATH", strIniPath)
   gstrPROEDIImpFP = GetSectionEntry("EDI_DEFAULTS", "PROEDI_IMPORT_FP", strIniPath)
   gstrPROEDIExpFP = GetSectionEntry("EDI_DEFAULTS", "PROEDI_EXPORT_FP", strIniPath)

   gstrRawEDIDir = GetSectionEntry("EDI_DEFAULTS", "PROEDI_RAW_EDIDIR", strIniPath)
   gstrRawEDIFile = GetSectionEntry("EDI_DEFAULTS", "PROEDI_RAW_EDIFILE", strIniPath)
   gstrRawEDIArch = GetSectionEntry("EDI_DEFAULTS", "PROEDI_RAW_EDIARCH_FP", strIniPath)


End Function


Public Function FindFileExten(ByVal strFileName As String, ByRef strFileEx As String)

   Dim lngPos As Long
   Dim iTotLen As Integer

   lngPos = InStr(strFileName, ".")
   iTotLen = Len(strFileName)

   If (lngPos > 0) Then
      strFileEx = Mid(strFileName, lngPos, ((iTotLen - lngPos) + 1))
   Else
      strFileEx = ""
   End If
End Function
