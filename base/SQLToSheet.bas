Sub ExportTableToSpreadsheet()
    ' Connect to the database
    Dim oContext As Object
    oContext = CreateUnoService("com.sun.star.sdb.DatabaseContext")
    Dim oDatabase As Object
    oDatabase = oContext.getByName("testdb")
    Dim oConnection As Object
    oConnection = oDatabase.getConnection("", "")
    
    ' Execute a SELECT statement to retrieve the table data
    Dim sSql As String
    sSql = "SELECT * FROM PARENTS JOIN CHILDREN on CHILDREN.PARENT=PARENTS.ID"
    Dim oStatement As Object
    oStatement = oConnection.createStatement()
    Dim oResultSet As Object
    oResultSet = oStatement.executeQuery(sSql)
    Dim oMetaData As Object
    oMetaData = oResultSet.getMetaData()
    
    ' Create a new spreadsheet document and write the data to it
    Dim oDesktop As Object
    oDesktop = createUnoService("com.sun.star.frame.Desktop")
    Dim oNewDoc As Object
    oNewDoc = oDesktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, Array())
    Dim oSheet As Object
    oSheet = oNewDoc.getSheets().getByIndex(0)
    
    ' Write the column headers to the first row of the sheet
    Dim iCol As Integer
    For iCol = 0 To oMetaData.getColumnCount() - 1
        oSheet.getCellByPosition(iCol, 0).setString(oMetaData.getColumnName(iCol + 1))
    Next iCol
    
    'write rows    
    Dim iRow As Integer
    iRow = 1
    While oResultSet.next()
        'Dim iCol As Integer
        iCol = 0
        While iCol < oResultSet.getMetaData().getColumnCount()
            oSheet.getCellByPosition(iCol, iRow).setValue(oResultSet.getString(iCol + 1))
            iCol = iCol + 1
        Wend
        iRow = iRow + 1
    Wend
    
    ' Close the result set and the connection
    oResultSet.close()
    oConnection.close()
End Sub
