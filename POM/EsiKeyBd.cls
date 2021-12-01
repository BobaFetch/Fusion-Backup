VERSION 1.0 CLASS
BEGIN
MultiUse = -1 'True
Persistable = 0 'NotPersistable
DataBindingBehavior = 0 'vbNone
DataSourceBehavior = 0 'vbNone
MTSTransactionMode = 0 'NotAnMTSObject
END
Attribute VB_Name = "EsiKeyBd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001, ES/2002) is the property of                     ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***

'See the UpdateTables prodecure for database revisions

Option Explicit

'//Class Module Added 6/16/98
'
'Sets formating for controls based on the "Tag" for
'each control.  Allows for less manual and more
'consistant coding of procedures.

'//Note: using these functions in control arrays
'        will produce an error (manual code).

'Tag  either text or combo (esTxt..., esCmb)
'---- -----------------------------------------
' 1 = value formated box
' 2 = upper/lower case formated box (default)
' 3 = upper case formatted box
' 4 = Date formatted textbox
' 5 = Time formatted textbox
' 6-7 reserved for boxes
' 8 = Locked (no edit) combobox
' 9   Multiline text Boxes
'10   Command button (future)
'11   option button  (future)

'esGotFocus for any text or combobox
'TextBoxes
Public WithEvents esTxtGotFocus As TextBox
Attribute esTxtGotFocus.VB_VarHelpID = -1
Public WithEvents esTxtKeyCase As TextBox '3
Attribute esTxtKeyCase.VB_VarHelpID = -1
Public WithEvents esTxtKeyCheck As TextBox '2 default
Attribute esTxtKeyCheck.VB_VarHelpID = -1
Public WithEvents esTxtKeyDate As TextBox '4
Attribute esTxtKeyDate.VB_VarHelpID = -1
Public WithEvents esTxtKeyMemo As TextBox '9
Attribute esTxtKeyMemo.VB_VarHelpID = -1
Public WithEvents esTxtKeyTime As TextBox '5
Attribute esTxtKeyTime.VB_VarHelpID = -1
Public WithEvents esTxtKeyValue As TextBox '1
Attribute esTxtKeyValue.VB_VarHelpID = -1
Public WithEvents esTxtKeyLock As TextBox '8
Attribute esTxtKeyLock.VB_VarHelpID = -1
Public WithEvents estxtKeyDown As TextBox
Attribute estxtKeyDown.VB_VarHelpID = -1

'ComboBoxes
Public WithEvents esCmbGotfocus As ComboBox
Attribute esCmbGotfocus.VB_VarHelpID = -1
Public WithEvents esCmbDropDown As ComboBox
Attribute esCmbDropDown.VB_VarHelpID = -1
Public WithEvents esCmbKeylock As ComboBox '8
Attribute esCmbKeylock.VB_VarHelpID = -1
Public WithEvents esCmbKeyCase As ComboBox '3
Attribute esCmbKeyCase.VB_VarHelpID = -1
Public WithEvents esCmbKeyDate As ComboBox '4
Attribute esCmbKeyDate.VB_VarHelpID = -1
Public WithEvents esCmbKeyValue As ComboBox '1
Attribute esCmbKeyValue.VB_VarHelpID = -1
Public WithEvents esCmbKeyCheck As ComboBox '2
Attribute esCmbKeyCheck.VB_VarHelpID = -1

'Masked Edit Boxes
Public WithEvents esMskGotFocus As MaskEdBox
Attribute esMskGotFocus.VB_VarHelpID = -1
Public WithEvents esMskKeyValue As MaskEdBox '1
Attribute esMskKeyValue.VB_VarHelpID = -1
Public WithEvents esMskKeyDown As MaskEdBox
Attribute esMskKeyDown.VB_VarHelpID = -1

'Show Printers
Public WithEvents esCmdClick As SSRibbon
Attribute esCmdClick.VB_VarHelpID = -1

Private Sub esCmdClick_Click(Value As Integer)
   If esCmdClick.Value = True Then
      'SysPrinters.Show
      esCmdClick.Value = False
   End If
   
End Sub

Private Sub esTxtKeyDown_KeyDown(KeyCode As Integer, Shift As Integer)
   'use in KeyDown
   'not for combo boxes or memo fields
   'to use vbKeyinsert you must have a label
   'name InsPanel (or something else)
   bUserAction = True
   If KeyCode = vbKeyDown Then
      SendKeys "{TAB}"
   Else
      If KeyCode = vbKeyUp Then SendKeys "+{TAB}"
   End If
   
End Sub

Private Sub esCmbGotfocus_GotFocus()
   Static a As Integer
   'Selects all of the text in fixed length text boxes
   'and combo boxes. Mostly gets rid of the blinking
   'num locks with VB5.0 (SP2) and SendKeys
   'Allow for the possibility that the control has been
   'disabled
   
   On Error Resume Next
   If Len(esCmbGotfocus) = 0 Then esCmbGotfocus = " "
   esCmbGotfocus.SelStart = 0
   esCmbGotfocus.SelLength = Len(esCmbGotfocus)
   
End Sub

Private Sub esTxtKeyMemo_KeyPress(KeyAscii As Integer)
   bUserAction = True
   If Not bInsertOn Then
      If KeyAscii > 13 Then esTxtKeyMemo.SelLength = 1
   End If
   
End Sub


Private Sub esTxtKeyValue_KeyPress(KeyAscii As Integer)
   'Allows only numbers, "-" and "." for value
   'fields like money or quantities
   'syntax in Keypress: KeyValue KeyAscii
   bUserAction = True
   If KeyAscii = 13 Then
      If bEnterAsTab Then
         KeyAscii = 0
         SendKeys "{TAB}"
      End If
   Else
      Select Case KeyAscii
         Case 8, 9, 32, 43 To 46, 48 To 57
         Case Else
            KeyAscii = 0
      End Select
      If Not bInsertOn Then
         If esTxtKeyValue.SelLength = 0 Then
            If KeyAscii > 13 Then esTxtKeyValue.SelLength = 1
         End If
      End If
   End If
   
End Sub



Private Sub esTxtGotFocus_GotFocus()
   Static a As Integer
   'Selects all of the text in fixed length text boxes
   'and combo boxes. Mostly gets rid of the blinking
   'num locks with VB5.0 (SP2) and SendKeys
   'Allow for the possibility that the control has been
   'disabled
   
   On Error Resume Next
   a = Len(esTxtGotFocus)
   If a = 0 Then esTxtGotFocus = " "
   esTxtGotFocus.SelStart = 0
   esTxtGotFocus.SelLength = Len(esTxtGotFocus)
   
End Sub

Private Sub esTxtKeyCase_KeyPress(KeyAscii As Integer)
   'All uppercase
   'syntax in Keypress Procedure KeyCase KeyAscii
   bUserAction = True
   If KeyAscii = 13 Then
      If bEnterAsTab Then
         KeyAscii = 0
         SendKeys "{TAB}"
      End If
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
      If Not bInsertOn Then
         If esTxtKeyCase.SelLength = 0 Then
            If KeyAscii > 13 Then esTxtKeyCase.SelLength = 1
         End If
      End If
   End If
   
End Sub

Private Sub esTxtKeyCheck_KeyPress(KeyAscii As Integer)
   'Check key for Enter key
   'syntax in Keypress Procedure KeyCheck KeyAscii
   bUserAction = True
   If KeyAscii = 13 Then
      If bEnterAsTab Then
         KeyAscii = 0
         SendKeys "{TAB}"
      End If
   Else
      If Not bInsertOn Then
         If esTxtKeyCheck.SelLength = 0 Then
            If KeyAscii > 13 Then esTxtKeyCheck.SelLength = 1
         End If
      End If
   End If
   
End Sub

Private Sub esTxtKeyDate_KeyPress(KeyAscii As Integer)
   'Changes ".", " " and "-" to "/" for dates
   'syntax in Keypress: KeyDate KeyAscii
   bUserAction = True
   If KeyAscii = 13 Then
      If bEnterAsTab Then
         KeyAscii = 0
         SendKeys "{TAB}"
      End If
   Else
      If KeyAscii = 8 Or KeyAscii = 9 Then Exit Sub
      Select Case KeyAscii
         Case Is < 43
            KeyAscii = 0
         Case Is > 57
            KeyAscii = 0
         Case Is = 45, 46
            KeyAscii = 47
      End Select
      If Not bInsertOn Then
         If esTxtKeyDate.SelLength = 0 Then
            If KeyAscii > 13 Then esTxtKeyDate.SelLength = 1
         End If
      End If
   End If
   
End Sub

Private Sub esCmbKeylock_KeyPress(KeyAscii As Integer)
   'Use only the Enter Key
   bUserAction = True
   If KeyAscii = 13 Then SendKeys "{TAB}"
   KeyAscii = 0
   
End Sub

Private Sub esCmbKeyCase_KeyPress(KeyAscii As Integer)
   'All uppercase
   'syntax in Keypress Procedure KeyCase KeyAscii
   If KeyAscii = 13 Then
      If bEnterAsTab Then
         KeyAscii = 0
         SendKeys "{TAB}"
      End If
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
      If Not bInsertOn Then
         If esCmbKeyCase.SelLength = 0 Then
            If KeyAscii > 13 Then esCmbKeyCase.SelLength = 1
         End If
      End If
      
   End If
   
End Sub

Private Sub esCmbKeyValue_Keypress(KeyAscii As Integer)
   'Allows only numbers, "-" and "." for value
   'fields like money or quantities
   'syntax in Keypress: KeyValue KeyAscii
   bUserAction = True
   bInsertOn = True
   If KeyAscii = 13 Then
      If bEnterAsTab Then
         KeyAscii = 0
         SendKeys "{TAB}"
      End If
   Else
      Select Case KeyAscii
         Case 8, 9, 32, 43 To 46, 48 To 57
         Case Else
            KeyAscii = 0
      End Select
      If Not bInsertOn Then
         If esCmbKeyValue.SelLength = 0 Then
            If KeyAscii > 13 Then esCmbKeyValue.SelLength = 1
         End If
      End If
   End If
   
End Sub

Private Sub esTxtKeyTime_KeyPress(KeyAscii As Integer)
   'Time field used in Keypress Event
   'Allows certain characters to be honored
   bUserAction = True
   If KeyAscii = 13 Then
      If bEnterAsTab Then
         KeyAscii = 0
         SendKeys "{TAB}"
      End If
   Else
      If bInsertOn Then
         If esTxtKeyTime.SelLength = 0 Then
            If KeyAscii > 13 Then esTxtKeyTime.SelLength = 1
         End If
      End If
      If KeyAscii = 8 Or KeyAscii = 9 Then Exit Sub
      If KeyAscii = 65 Then KeyAscii = 97
      If KeyAscii = 80 Then KeyAscii = 112
      If KeyAscii = 45 Then KeyAscii = 58
      If KeyAscii = 46 Then KeyAscii = 58
      If KeyAscii = 58 Or KeyAscii = 97 Or KeyAscii = 112 Then Exit Sub
      Select Case KeyAscii
         Case Is < 43, 47, Is > 57
            KeyAscii = 0
      End Select
   End If
   
End Sub


Private Sub Documentation()
   '//Paste the following in the Declarations of
   'the form:
   
   'Private txtKeyPress() As New EsiKeyBd
   'Private txtGotFocus() As New EsiKeyBd
   'Private txtKeyDown()  As New EsiKeyBd
   
   '//Create a new sub for the form and add the code
   'Call the procedure from Form_Load after any Move
   'Commands
   
   '//See the 2nd example for small Control count forms
   
   '//First example:
   'Private Sub FormatControls()
   'See class definitions for "Tags"
   'Tags may be coded here or in the
   'Propertys Window for each control
   
   '    Dim bByte As Byte
   '    Dim i As Integer
   '    Dim a As Integer
   '    Dim b As Integer
   '    Dim c As Integer
   '    a = -1
   '    b = -1
   '
   '     ' //Need the following in case of a untrapped Control Array.
   '     ' Manual Code those from Module Procedures.
   '        On Error Resume Next
   '        For i = 0 To Controls.Count - 1
   '           'Part of an Array or label (z1(n))?
   '            c = Controls(i).Index
   '           If Err > 0 And (TypeOf Controls(i) Is textBox Or _
   '              TypeOf Controls(i) Is ComboBox Or TypeOf Controls(i) Is MaskEdBox) Then
   '                Err = 0
   '                a = a + 1
   '                ReDim Preserve txtKeyPress(a) As New EsiKeyBd
   '                If Controls(i).Tag <> "9" Then
   '                    b = b + 1
   '                    ReDim Preserve txtGotFocus(b) As New EsiKeyBd
   '                    ReDim Preserve txtKeyDown(b) As New EsiKeyBd
   '                End If
   '                If TypeOf Controls(i) Is MaskEdBox Then
   '                    Set txtGotFocus(b).esMskGotFocus = Controls(i)
   '                    Set txtKeyDown(b).esMskKeyDown = Controls(i)
   '                    Set txtKeyPress(a).esMskKeyValue = Controls(i)
   '                End If
   '                If TypeOf Controls(i) Is textBox Then
   '                  bByte = True
   '                  Select Case Val(Controls(i).Tag)
   '                      Case 1
   '                          Set txtKeyPress(a).esTxtKeyValue = Controls(i)
   '                      Case 3
   '                          Set txtKeyPress(a).esTxtKeyCase = Controls(i)
   '                      Case 4
   '                          Set txtKeyPress(a).esTxtKeyDate = Controls(i)
   '                      Case 5
   '                          Set txtKeyPress(a).esTxtKeyTime = Controls(i)
   '                      Case 9
   '                          Set txtKeyPress(a).esTxtKeyMemo = Controls(i)
   '                          bByte = False
   '                      Case Else
   '                          Set txtKeyPress(a).esTxtKeyCheck = Controls(i)
   '                  End Select
   '                  If bByte Then
   '                      Set txtGotFocus(b).esTxtGotfocus = Controls(i)
   '                      Set txtKeyDown(b).esTxtKeyDown = Controls(i)
   '                  End If
   '               Else
   '                  If TypeOf Controls(i) Is ComboBox Then
   '                    Set txtGotFocus(b).esCmbGotfocus = Controls(i)
   '                    Select Case Val(Controls(i).Tag)
   '                      Case 1
   '                          Set txtKeyPress(a).esCmbKeyValue = Controls(i)
   '                      Case 4
   '                          Set txtKeyPress(a).esCmbKeyDate = Controls(i)
   '                      Case 8
   '                          Set txtKeyPress(a).esCmbKeylock = Controls(i)
   '                           Controls(i).ForeColor = ES_BLUE
   '                      Case Else
   '                          Set txtKeyPress(a).esCmbKeyCase = Controls(i)
   '                    End Select
   '                  End If
   '              End If
   '            End If
   '        Next
   
   'End Sub
   
   
End Sub



Private Sub esTxtKeyLock_KeyPress(KeyAscii As Integer)
   'Use only the Enter Key
   bUserAction = True
   If KeyAscii = 13 Then SendKeys "{TAB}"
   If KeyAscii > 9 Then KeyAscii = 0
   
End Sub


Private Sub esMskGotFocus_GotFocus()
   Static a As Integer
   'Selects all of the text in fixed length text boxes
   'and combo boxes. Mostly gets rid of the blinking
   'num locks with VB5.0 (SP2) and SendKeys
   'Allow for the possibility that the control has been
   'disabled
   
   On Error Resume Next
   If Len(esMskGotFocus) = 0 Then esMskGotFocus = " "
   esMskGotFocus.SelStart = 0
   esMskGotFocus.SelLength = Len(esMskGotFocus)
   
End Sub

Private Sub esMskKeyValue_Keypress(KeyAscii As Integer)
   'Allows only numbers, "-" and "." for value
   'fields like money or quantities
   'syntax in Keypress: KeyValue KeyAscii
   bUserAction = True
   bInsertOn = True
   If KeyAscii = 13 Then
      If bEnterAsTab Then
         KeyAscii = 0
         SendKeys "{TAB}"
      End If
   Else
      Select Case KeyAscii
         Case 8, 9, 32, 43 To 46, 48 To 57
         Case Else
            KeyAscii = 0
      End Select
      If Not bInsertOn Then
         If esMskKeyValue.SelLength = 0 Then
            If KeyAscii > 13 Then esMskKeyValue.SelLength = 1
         End If
      End If
   End If
   
End Sub

Private Sub esMskKeyDown_KeyDown(KeyCode As Integer, Shift As Integer)
   'use in KeyDown
   'not for combo boxes or memo fields
   'to use vbKeyinsert you must have a label
   'name InsPanel (or something else)
   bUserAction = True
   If KeyCode = vbKeyDown Then
      SendKeys "{TAB}"
   Else
      If KeyCode = vbKeyUp Then SendKeys "+{TAB}"
   End If
   
End Sub

Private Sub esCmbKeyDate_Keypress(KeyAscii As Integer)
   'Changes ".", " " and "-" to "/" for dates
   'syntax in Keypress: KeyDate KeyAscii
   bUserAction = True
   If KeyAscii = 13 Then
      If bEnterAsTab Then
         KeyAscii = 0
         SendKeys "{TAB}"
      End If
   Else
      If KeyAscii = 8 Then Exit Sub
      Select Case KeyAscii
         Case Is < 43
            KeyAscii = 0
         Case Is > 57
            KeyAscii = 0
         Case Is = 45, 46
            KeyAscii = 47
      End Select
      If Not bInsertOn Then
         If esCmbKeyDate.SelLength = 0 Then
            If KeyAscii > 13 Then esCmbKeyDate.SelLength = 1
         End If
      End If
   End If
   
End Sub


Private Sub esCmbKeyCheck_Keypress(KeyAscii As Integer)
   'Check key for Enter key
   'syntax in Keypress Procedure KeyCheck KeyAscii
   bUserAction = True
   If KeyAscii = 13 Then
      If bEnterAsTab Then
         KeyAscii = 0
         SendKeys "{TAB}"
      End If
   Else
      If Not bInsertOn Then
         If esCmbKeyCheck.SelLength = 0 Then
            If KeyAscii > 13 Then esCmbKeyCheck.SelLength = 1
         End If
      End If
   End If
   
End Sub

Private Sub esCmbDropDown_DropDown()
   ShowCalendar MdiSect.ActiveForm
   
End Sub
