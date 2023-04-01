Sub PrintParagraphStyles()
    Dim oDoc As Object
    Dim oParas As Object
    Dim oPara As Object
    
    'Get the current document
    oDoc = ThisComponent
    
    'Get all the paragraphs in the document
    oParas = oDoc.Text.createEnumeration()
    
    'Loop over the paragraphs and print their style names
    Do While oParas.hasMoreElements()
        oPara = oParas.nextElement()
        If oPara.supportsService("com.sun.star.text.Paragraph") Then
            Print oPara.getString()
        End If
    Loop
End Sub
