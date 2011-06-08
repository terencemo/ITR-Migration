Rem Attribute VBA_ModuleType=VBADocumentModule
Option VBASupport 1
Private Sub Worksheet_Activate()
MsgBox "To Save the XML which, you have to click on the SAVE XML button on this sheet.", vbInformation
End Sub

Private Sub Worksheet_Deactivate()
Sheet4.Visible = xlSheetHidden
End Sub
