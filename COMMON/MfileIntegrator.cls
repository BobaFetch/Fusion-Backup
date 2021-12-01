VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MfileIntegrator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private Const REG_SZ As Long = 1 'REG_SZ represents a fixed-length text string.
Private Const REG_DWORD As Long = 4 'REG_DWORD represents data by a number that is 4 bytes long.
Private Const HKEY_LOCAL_MACHINE = &H80000002 'Contains configuration information particular to the computer (for any user).
Private Const ERROR_NONE = 0
Private Const KEY_QUERY_VALUE = &H1 'Required to query the values of a registry key.
Const MAX_PATH = 260

'Close Registry Key
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
'Open Registry Key
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
'Query a String Value
Private Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
'Query a Long Value
Private Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long
'Query a NULL Value
Private Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long
'Get Windows Temp folder path
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias _
"GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
ByVal lpKeyName As Any, ByVal lpDefault As String, _
ByVal lpReturnedString As String, ByVal nSize As Long, _
ByVal lpFileName As String) As Long

Public Enum eMFDataType
   MFDatatypeText = 1
   MFDatatypeInteger = 2
   MFDatatypeFloating = 3
   MFDatatypeDate = 5
   MFDatatypeBoolean = 8
End Enum


Private sDocLokInstallDirectory As String
Private objDoc As DOMDocument
Private objRootNode As IXMLDOMNode
Private objNodeVault As IXMLDOMNode
Private objNodeCom As IXMLDOMNode
Private objNodeMDoc As IXMLDOMNode

Private sXMLFileName As String

Public gsVaultName As String
Public gsMOPOClassID As String
Public gsMOClassID As String
Public gsCRClassID As String
Public gsVndInvClassID As String
Public gsDOCClassID As String

Public gsMfileInt As String
Public gsMFileWorkDir As String
Public gsXmlFileName As String

Dim bAutoScan As Integer


Private Sub Class_Initialize()
    'sDocLokInstallDirectory = QueryValue("Software\aidmc", "InstallDir")
    Dim strIniPath As String
    strIniPath = App.Path & "\" & "MFileInit.ini"
    
    MFileSetting strIniPath
    
    GetDisAutoScanFlag
       
    
    'MsgBox gstrRegPath
    'MsgBox gstrRegKeyName
    
    If (gsVaultName = "") Then
      MsgBox "M-File Vault is Empty.", vbInformation
    End If
End Sub


Private Function GetTmpPath()
    Dim sFolder As String ' Name of the folder
    Dim lRet As Long ' Return Value

    sFolder = String(MAX_PATH, 0)
    lRet = GetTempPath(MAX_PATH, sFolder)

    If lRet <> 0 Then
        GetTmpPath = Left(sFolder, InStr(sFolder, _
        Chr(0)) - 1)
    Else
        GetTmpPath = vbNullString
    End If
End Function


Private Function QueryValue(sKeyName As String, sValueName As String) As String
       
  Dim lRetVal As Long         'result of the API functions
  Dim hKey As Long         'handle of opened key
  Dim vValue As Variant      'setting of queried value

  lRetVal = RegOpenKeyEx(HKEY_LOCAL_MACHINE, sKeyName, 0, KEY_QUERY_VALUE, hKey) 'Open Key to Query a value
  If ERROR_NONE = lRetVal Then lRetVal = QueryValueEx(hKey, sValueName, vValue) 'Query (determine) the value stored
  RegCloseKey (hKey) 'Close the Key
  If lRetVal = ERROR_NONE Then QueryValue = vValue Else QueryValue = ""
End Function


Function QueryValueEx(ByVal lhKey As Long, ByVal szValueName As String, vValue As Variant) As Long
       Dim Data As Long
       Dim retval As Long 'Return value of RegQuery functions
       Dim lType As Long 'Determine data type of present data
       Dim lValue As Long 'Long value
       Dim sValue As String 'String value

       On Error GoTo QueryValueExError

       ' Determine the size and type of data to be read
       retval = RegQueryValueExNULL(lhKey, szValueName, 0&, lType, 0&, Data)
       
       If retval <> ERROR_NONE Then Error 5

       Select Case lType
           ' Determine strings
           Case REG_SZ:
               sValue = String(Data, 0)

               retval = RegQueryValueExString(lhKey, szValueName, 0&, lType, sValue, Data)
               
               If retval = ERROR_NONE Then
                   vValue = Left$(sValue, Data - 1)
               Else
                   vValue = Empty
               End If
               
           ' Determine DWORDS
           Case REG_DWORD:
               retval = RegQueryValueExLong(lhKey, szValueName, 0&, lType, lValue, Data)
               
               If retval = ERROR_NONE Then vValue = lValue
           
           Case Else
               'all other data types not supported
               retval = -1
       End Select
    
QueryValueExError:
       QueryValueEx = retval
       Exit Function

End Function

Private Sub CreateProcessingInstruction(XMLDoc As DOMDocument, InstrName As String, InstrValue As String)
    Dim objNode As IXMLDOMProcessingInstruction

    Set objNode = XMLDoc.CreateProcessingInstruction(InstrName, InstrValue)
    XMLDoc.appendChild objNode
    Set objNode = Nothing
End Sub

Private Sub CreateComment(XMLDoc As DOMDocument, ByVal CommentText As String)
    Dim objComment As IXMLDOMComment
    
    Set objComment = XMLDoc.CreateComment(CommentText)
    XMLDoc.appendChild objComment
    
    Set objComment = Nothing
End Sub

Private Function CreateNode(XMLDoc As DOMDocument, ParentNode As IXMLDOMNode, NodeName As String, Optional NodeValue As String = "") As IXMLDOMNode
    Dim objNode As IXMLDOMNode
    'Create the node
    Set objNode = XMLDoc.createElement(NodeName)
    
    'Add the text if it has any
    If Len(Trim(NodeValue)) > 0 Then
        objNode.Text = NodeValue
    End If
    
    'Add the node to the document
    If ParentNode Is Nothing Then
        XMLDoc.appendChild objNode
    Else
        ParentNode.appendChild objNode
    End If
    
    Set CreateNode = objNode
    
    Set objNode = Nothing
End Function

Private Sub CreateAttribute(XMLDoc As DOMDocument, oNode As IXMLDOMNode, AttribName As String, AttribValue As String)
Dim oElement As IXMLDOMElement

    Set oElement = oNode
    oElement.setAttribute AttribName, AttribValue
    
    Set oElement = Nothing

End Sub


Public Sub OpenXMLFile(ByVal sDirective As String, ByVal sCollection As String, _
   ByVal sContainer As String, ByVal sClassID As String, Optional sOptions As String = "", _
   Optional sFileName As String = "", Optional sFilePath As String = "")

    
   If gsMFileWorkDir <> "" And gsXmlFileName <> "" Then
       'use temp windows file directory and name
       sXMLFileName = gsMFileWorkDir & gsXmlFileName
   Else
      MsgBox ("DocLok Integrator.exe not found")
      Exit Sub
   End If
    'Check if the log exists. If it does open it otherwise create it
    If Len(Dir(sXMLFileName)) > 0 Then Kill sXMLFileName

    Set objDoc = New DOMDocument
    
    'Create the processing instructions
    CreateProcessingInstruction objDoc, "xml", "version='1.0'"
    
    'Create the Root Node
    Set objRootNode = CreateNode(objDoc, Nothing, "FusionIntegrator")
    CreateAttribute objDoc, objRootNode, "xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance"
    CreateAttribute objDoc, objRootNode, "xmlns:xsd", "http://www.w3.org/2001/XMLSchema"
    CreateAttribute objDoc, objRootNode, "xmlns", "http://www.cpandl.com"
    
    'Start Vault xml
    Set objNodeVault = CreateNode(objDoc, objRootNode, "MVault")
    CreateAttribute objDoc, objNodeVault, "VaultName", gsVaultName
    CreateAttribute objDoc, objNodeVault, "DOCClassID", sClassID
    
    ' Command xml
    Set objNodeCom = CreateNode(objDoc, objRootNode, "MCommand")
    CreateAttribute objDoc, objNodeCom, "Directive", sDirective
    CreateAttribute objDoc, objNodeCom, "ConfigurationReferenceID", "1"
    
    ' Document
    Set objNodeMDoc = CreateNode(objDoc, objRootNode, "MDocument")
    CreateAttribute objDoc, objNodeMDoc, "KeyObject", sCollection
    CreateAttribute objDoc, objNodeMDoc, "Container", sContainer
    CreateAttribute objDoc, objNodeMDoc, "Name", sFileName
    
End Sub

Public Sub AddXMLMetaData(ByVal sMetaDataName As String, ByVal sMetaDataID As Long, _
                              ByVal sMetaData As String, Optional eMType As eMFDataType = eMFDataType.MFDatatypeText)

   Dim objMetaNode As IXMLDOMNode
   Dim objMetaData As IXMLDOMNode
   
   Set objMetaNode = CreateNode(objDoc, objNodeMDoc, "MMetaData")
   CreateAttribute objDoc, objMetaNode, "Name", sMetaDataName
   CreateAttribute objDoc, objMetaNode, "Id", CStr(sMetaDataID)
   CreateAttribute objDoc, objMetaNode, "type", CStr(eMType)
   
   Set objMetaData = CreateNode(objDoc, objMetaNode, "data", sMetaData)
   
   Set objMetaData = Nothing
   Set objMetaNode = Nothing
   
End Sub


Public Sub CloseXMLFile()
    objDoc.save sXMLFileName
    
    Set objNodeMDoc = Nothing
    Set objNodeCom = Nothing
    Set objNodeVault = Nothing
    Set objRootNode = Nothing
    Set objDoc = Nothing

End Sub


Public Function SendXMLFileToMFile() As Boolean
   On Error GoTo modErr1
   Err = 0
   
   
   If (gsMfileInt <> "") Then
      Dim strFullPath As String
      'strFullPath = App.Path & "\" & gsMfileInt & " " & gsMFileWorkDir & gsXmlFileName
      strFullPath = App.Path & "\" & gsMfileInt & " " & gsXmlFileName & " " & CStr(bAutoScan)
      Err.Clear
      
      ' if auto scanning, hide FusionScan.  Otherweise show it.
      If bAutoScan = 0 Then
         Shell strFullPath, vbNormalFocus
      Else
         Shell strFullPath, vbHide
      End If
      
   Else
      MsgBox ("DocLok Integrator.exe not found")
   End If
   If Err > 0 Then SendXMLFileToMFile = False Else SendXMLFileToMFile = True
   
   Exit Function
   
modErr1:
   sProcName = "GetSectionEntry"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
End Function

Private Function GetSectionEntry(ByVal strSectionName As String, ByVal strEntry As String, ByVal strIniPath As String) As String


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


Private Function MFileSetting(strIniPath As String)

   gsVaultName = GetSectionEntry("MFILES_DEFAULTS", "VAULTNAME", strIniPath)
   gsMOPOClassID = GetSectionEntry("MFILES_DEFAULTS", "MOPOCLASSID", strIniPath)
   gsMOClassID = GetSectionEntry("MFILES_DEFAULTS", "MOCLASSID", strIniPath)
   gsCRClassID = GetSectionEntry("MFILES_DEFAULTS", "CRCLASSID", strIniPath)
   gsVndInvClassID = GetSectionEntry("MFILES_DEFAULTS", "VNDINVCLASSID", strIniPath)
   gsDOCClassID = GetSectionEntry("MFILES_DEFAULTS", "DOCCLASSID", strIniPath)
   
   gsMfileInt = GetSectionEntry("MFILES_DEFAULTS", "MFILEINTEGRATOR", strIniPath)
   gsXmlFileName = GetSectionEntry("MFILES_DEFAULTS", "MINTFILENAME", strIniPath)
   gsMFileWorkDir = GetSectionEntry("MFILES_DEFAULTS", "MFILEWORKINGDIR", strIniPath)
   
    'create xml directory  C:\MyWorkFiles files if it does not exist
    If Dir(gsMFileWorkDir, vbDirectory) = "" Then
      MkDir gsMFileWorkDir
    End If
    

End Function


Private Function GetDisAutoScanFlag()

   bAutoScan = 1
   Dim RdoScan As ADODB.Recordset
   sSql = "SELECT ISNULL(CODISAUTOSCAN, 0) CODISAUTOSCAN FROM ComnTable WHERE COREF=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoScan, ES_FORWARD)
   If bSqlRows Then
      With RdoScan
         If (CInt(!CODISAUTOSCAN) = 1) Then
            bAutoScan = 0
         End If
         ClearResultSet RdoScan
      End With
   Else
      bAutoScan = 1
   End If
   Set RdoScan = Nothing


End Function

