global const DBName = "New Database"
global const SQLInCol = 3

Sub LoadSheetToDB()
    ' Open a file picker to select the spreadsheet file
    Dim oDoc As Object
    Dim oFilePicker As Object
    oFilePicker = createUnoService("com.sun.star.ui.dialogs.FilePicker")
    If oFilePicker.execute() <> 1 Then
        Exit Sub
    End If
    oDoc = StarDesktop.loadComponentFromURL(oFilePicker.files(0), "_blank", 0, Array())
    Dim oSheet As Object
    oSheet = oDoc.Sheets.getByName("Sheet1")
    
    ' Connect to the database
    Dim oContext As Object
    oContext = CreateUnoService("com.sun.star.sdb.DatabaseContext")
    Dim oDatabase As Object
    oDatabase = oContext.getByName(DBName)
    Conn = oDatabase.getConnection("","") 
    Dim Stmt As Object
	Dim Result As Object
	dim strSQL as String
    
    ' Load the data row-wise
	i = 1
	oCell = oSheet.getCellByPosition(SQLInCol,i)
	while oCell.Type <> com.sun.star.table.CellContentType.EMPTY
		strSQL = oCell.String
		Stmt = Conn.createStatement()
		Result = Stmt.executeQuery(strSQL)
		
	    i = i+1
	    oCell = oSheet.getCellByPosition(SQLInCol,i)
	wend
    
    ' Close the spreadsheet
    oDoc.Close(True)
End Sub
