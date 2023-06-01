Sub CreateTextBox
    Dim oDoc As Object
    Dim oShape As Object
    Dim sText As String
    Dim oSize As New com.sun.star.awt.Size
	Dim oPosition As New com.sun.star.awt.Point
    
    sText = "This is the text"
    
    oDoc = ThisComponent
	oShape = thisComponent.createInstance("com.sun.star.drawing.TextShape")

	'size of text box
	oSize.Width  = 1700                                          
	oSize.Height = 800                                         
	'position of each text box relative to the size of the page
	oPosition.X = 2000
	oPosition.Y = 2000
	
	oShape.Size             = oSize
	oShape.Position         = oPosition	
	oShape.rotateAngle      = 0
	oShape.AnchorType       = com.sun.star.text.TextContentAnchorType.AT_PAGE
	oShape.AnchorPosition.X = 0
	oShape.AnchorPosition.Y = 0
	oShape.AnchorPageNo		= i
	oShape.Name = "TextBox_"&sText

	oDoc.drawPage.add(oShape)
	
	oShape.String = sText
End Sub
