Sub WriteHelloWorld()

    'Open the file
    Dim oDoc As Object
    Dim oFilePicker As Object
    oFilePicker = createUnoService("com.sun.star.ui.dialogs.FilePicker")
    If oFilePicker.execute() <> 1 Then
        Exit Sub
    End If
    oDoc = StarDesktop.loadComponentFromURL(oFilePicker.files(0), "_blank", 0, Array())
    
    'Write "Hello World"
    Dim oText As Object
    oText = oDoc.Text
    oText.insertString(oText.End, "Hello World", False)
    
    'Save the file
    oDoc.store()
    oDoc.close(True)
    
End Sub
