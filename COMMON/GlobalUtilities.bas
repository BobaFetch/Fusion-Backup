Attribute VB_Name = "GlobalUtilities"
'This module contains utilities shared by the various ES2000 modules
'and Esi2000.exe
'Most code that is shared by ES2000 modules but not Esi2000.exe is in ESIPROJ.bas

Option Explicit

Public Function Debugging() As Boolean
   If InStr(1, Command, "/debug", vbTextCompare) Then
      Debugging = True
   Else
      Debugging = False
   End If
End Function

Public Function RunningBeta() As Boolean
   'determine whether running beta featurs
   'used initially to replace ES_CUSTOM = "PROPLA" with RunningBeta
   If InStr(1, Command, "/beta", vbTextCompare) Then
      RunningBeta = True
   Else
      RunningBeta = False
   End If
End Function


'BBS Made this function to try and centralize the reading of the preferences
'BE WARNED That if, for example, you are reading a bit or true/false field
'That this fuction will return false (which should be ok) even if the value
'of the preference field is null
Public Function GetPreferenceValue(strFldNme As String, Optional CompanyPreference As Boolean = False) As String
    Dim RdoPref As ADODB.Recordset
    
    On Error Resume Next
    GetPreferenceValue = ""
    If CompanyPreference Then
        sSql = "SELECT " & strFldNme & " AS PrefFld FROM ComnTable WHERE COREF=1"
    Else
        sSql = "SELECT " & strFldNme & " AS PrefFld FROM Preferences WHERE PreRecord=1"
    End If
    bSqlRows = clsADOCon.GetDataSet(sSql, RdoPref, ES_FORWARD)
    
    If bSqlRows Then
        GetPreferenceValue = "" & RdoPref!PrefFld
        ClearResultSet RdoPref
    End If
    
    Set RdoPref = Nothing
End Function



Public Function InsertAuditEntry(strTblName As String, strSQL As String)

   On Error Resume Next
   Dim tmpsql As String
   tmpsql = strSQL
   
   If (TableExists(strTblName)) Then
      If (tmpsql <> "") Then
         clsADOCon.ExecuteSql sSql ' rdExecDirect
      End If
   End If
End Function

'Public Function GetAuditTransID(strTblName As String) As Long
'
'   On Error Resume Next
'   Dim RdoTransID As ADODB.Recordset
'
'   GetAuditTransID = -1
'   If (TableExists(strTblName)) Then
'
'      sSql = "SELECT (MAX(ADT_TRANS_ID) + 1)as LastTransID FROM " & strTblName & ""
'      bSqlRows = clsADOCon.GetDataSet(sSql, RdoTransID, ES_FORWARD)
'      If bSqlRows Then
'          GetAuditTransID = IIf(IsNull(RdoTransID!LastTransID), 1, RdoTransID!LastTransID)
'          ClearResultSet RdoTransID
'      End If
'
'       Set RdoTransID = Nothing
'   End If
'End Function

Public Function GetPrinterPort(devPrinter As String, devDriver As String, devPort As String) As Byte
   Dim SysPrinter As Printer
   For Each SysPrinter In Printers
      If Trim(SysPrinter.DeviceName) = devPrinter Then
         devDriver = SysPrinter.DriverName
         devPort = SysPrinter.Port
         Exit For
      End If
   Next
   
End Function

