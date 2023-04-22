' load files from cells in a spreadsheet and run a macro against them
Sub processFilesInColumnA
    
    dim oSheet as Object
    dim filename as String
    dim oCell as Object
    dim i as Integer
	
    ' Get the active sheet
    oSheet = ThisComponent.getCurrentController().getActiveSheet()

	' iterate down column A getting the filepaths
	i = 0
	oCell = oSheet.getCellByPosition(0,i)
	while oCell.Type <> com.sun.star.table.CellContentType.EMPTY
		filename = ConvertToURL(oCell.String)
		processFile(filename)
		
	    i = i+1
	    oCell = oSheet.getCellByPosition(0,i)
	wend
End Sub

Function processFile(filename As String)
	dim oDoc as Object
	
    oDoc = StarDesktop.loadComponentFromURL(filename, "_blank", 0, Array())
    
    'Write "Hello World"
    Dim oText As Object
    oText = oDoc.Text
    oText.insertString(oText.End, "Hello World", False)
    
    'Save the file
    oDoc.store()
    oDoc.close(True)
End Function
