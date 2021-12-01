VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ComboExpander"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' The ComboBox we are expanding.
Private WithEvents m_ComboBox As ComboBox
Attribute m_ComboBox.VB_VarHelpID = -1

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Const CB_FINDSTRINGEXACT = &H158
Private Const CB_FINDSTRING = &H14C
Private Const CB_SETCURSEL = &H14E

' Prepare to expand the ComboBox.
Public Sub Initialize(ByVal cbo As ComboBox)
    Set m_ComboBox = cbo
End Sub
' The user typed this key.
Private Sub ProcessKey(ByRef key_code As Integer)
Dim new_text As String
Dim row_found As Long

    ' See what the ComboBox's new text will be.
    new_text = NewComboText(key_code)

    ' Find the first row that matches this string.
    ' See if the character is Backspace or Delete.
    If key_code = vbKeyBack Or key_code = vbKeyDelete Then
        ' This is Backspace or Delete.
        ' Require an exact match.
        row_found = SendMessageString(m_ComboBox.hwnd, _
            CB_FINDSTRINGEXACT, 0, new_text)
    Else
        ' This is not Backspace or Delete.
        ' Require any match.
        row_found = SendMessageString(m_ComboBox.hwnd, _
            CB_FINDSTRING, 0, new_text)
    End If

    ' See if we found a row.
    If row_found <> -1 Then
        ' We found a row. Select the row.
        SendMessage m_ComboBox.hwnd, _
            CB_SETCURSEL, row_found, ByVal 0&

        ' Select the text.
        m_ComboBox.SelStart = Len(new_text)
        m_ComboBox.SelLength = Len(m_ComboBox.Text) - Len(new_text)

        ' Cancel the input character so the ComboBox
        ' doesn't try to use it.
        key_code = 0
    End If
End Sub
' Determine the text's new value based on the
' currently selected text and the character pressed.
Private Function NewComboText(ByVal key_asc As Integer) As String
Dim txt_left As String
Dim txt_selected As String
Dim txt_right As String
Dim result As String

    ' Get the pieces of the ComboBox's text.
    txt_left = Left$(m_ComboBox.Text, m_ComboBox.SelStart)
    txt_selected = Mid$(m_ComboBox.Text, m_ComboBox.SelStart + 1, m_ComboBox.SelLength)
    txt_right = Mid$(m_ComboBox.Text, m_ComboBox.SelStart + m_ComboBox.SelLength + 1)

    ' Take action based on the character.
    Select Case key_asc
        Case vbKeyBack  ' Backspace.
            ' See if any text is selected.
            If Len(txt_selected) > 0 Then
                ' Text is selected. Delete it.
                result = txt_left & txt_right
            Else
                ' No text is selected.
                ' Delete one character to the left.
                If Len(txt_left) > 0 Then
                    txt_left = Left$(txt_left, Len(txt_left) - 1)
                End If
                result = txt_left & txt_right
            End If

        Case vbKeyDelete    ' Delete.
            ' See if any text is selected.
            If Len(txt_selected) > 0 Then
                ' Text is selected. Delete it.
                result = txt_left & txt_right
            Else
                ' No text is selected.
                ' Delete one character to the right.
                txt_right = Mid$(txt_right, 2)
                result = txt_left & txt_right
            End If

        Case Else           ' Some other character.
            ' Replace the selected text with the character.
            result = txt_left & Chr$(key_asc) & txt_right
    End Select

    NewComboText = result
End Function
' Watch for the Delete key.
Private Sub m_ComboBox_KeyDown(KeyCode As Integer, Shift As Integer)
    ' If it's a Delete key, process it.
    If KeyCode = vbKeyDelete Then ProcessKey KeyCode
End Sub

' Process printing keys and Backspace.
Private Sub m_ComboBox_KeyPress(KeyAscii As Integer)
Const ASC_SPACE As Integer = 32
Const ASC_TILDE As Integer = 126

    ' Ignore the key if it is non-printing
    ' and not Backspace.
    If (KeyAscii < ASC_SPACE Or KeyAscii > ASC_TILDE) _
        And KeyAscii <> vbKeyBack _
    Then
        ' The key is non-printing and not Backspace.
        Exit Sub
    End If

    ' Process the key.
    ProcessKey KeyAscii
End Sub
