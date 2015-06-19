Attribute VB_Name = "GenericMessageBoxFns"
Option Explicit

Public BasicMessageBoxReturnValue As Long

Public Function DisplayBasicMessageBox(widthTwips As Long, heightTwips As Long, theMessage As String, buttonTexts As Variant, Optional textAlignment As AlignmentConstants = vbCenter) As Long
    Load GenericMessageBox
    GenericMessageBox.CustomizeForm widthTwips, heightTwips, theMessage, buttonTexts, textAlignment
    GenericMessageBox.Show MODAL
    'at this point, GenericMessageBoxFns.BasicMessageBoxReturnValue gets populated
    ' with the index of the button pushed, so just return that
    DisplayBasicMessageBox = BasicMessageBoxReturnValue
End Function

