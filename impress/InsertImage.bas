Sub InsertImageWithLink
    Dim oDrawPage As Object
    Dim oImage As Object
    Dim oFilePicker As Object
    Dim sFileName As String
    Dim oController As Object
    Dim oSize As New com.sun.star.awt.Size

    
    oController = ThisComponent.getCurrentController()
    oDrawPage = oController.currentPage
    
    oFilePicker = createUnoService("com.sun.star.ui.dialogs.FilePicker")
    oFilePicker.setTitle("Select Image")
    oFilePicker.setMultiSelectionMode(False)
    If oFilePicker.execute() = 1 Then
        sFileName = oFilePicker.getFiles()(0)
        oImage = ThisComponent.createInstance("com.sun.star.drawing.GraphicObjectShape")
        oImage.GraphicURL = ConvertToURL(sFileName)
        oImage.Position = CreateUnoStruct("com.sun.star.awt.Point", 0, 0)
    	oSize.Width = oDrawPage.Width
    	oSize.Height = oDrawPage.Height
    	oImage.setSize(oSize)
        oDrawPage.add(oImage)
    End If
End Sub
