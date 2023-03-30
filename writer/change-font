Sub InsertHelloWorld()
    Dim oDoc As Object
    Dim oText As Object
    Dim oCursor As Object
    
    'Create a new document
    oDoc = StarDesktop.loadComponentFromURL("private:factory/swriter", "_blank", 0, Array())
    oText = oDoc.Text
    oCursor = oText.createTextCursor()
    
    'Insert first paragraph with "Hello World" in Arial font
    oCursor.CharWeight = com.sun.star.awt.FontWeight.BOLD
    oCursor.CharFontName = "Arial"
    oText.insertString(oCursor, "Hello World", False)
    
    'Insert second paragraph with "Hello World" in Times New Roman font
    oText.insertControlCharacter(oCursor, com.sun.star.text.ControlCharacter.PARAGRAPH_BREAK, False)
    oCursor.CharWeight = com.sun.star.awt.FontWeight.NORMAL
    oCursor.CharFontName = "Times New Roman"
    oText.insertString(oCursor, "Hello World", False)

End Sub
