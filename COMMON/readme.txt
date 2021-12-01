Changes -

Find some forms that you may be using included. The Form1 only contains
updated bitmaps and those should be pasted on the SideBar Panel
of your MdiSect.

The only code that you need to add or change is:
            Case "2005"
                Image3(2).Picture = img05.Picture
                Image3(4).Picture = img05.Picture

To the MdiSect.FormLoad in the Select Statement.

The bitmaps will automatically be updated in the SetDiaPos procedure.
They will also be aligned and resized.

A color Constant has been added:
Public Const ES_ViewBackColor = &HE0FFFF   '11/17/04 RGB(255, 253, 223)

This is used in the Form_Initialize to provide a constant color background for
LookUps that is visible when the user has the default 'windows App Background.
Suggest you use the same scheme in your customs.

There are (2) new procedures that I am using to consolidate the code and lower
the footprint in filling ComboBoxes:

LoadComboBox
LoadNumComboBox

I would not spend a great deal of time updating to those, but they are handy and do
make the FillCombo procedure much smaller and more readable.
