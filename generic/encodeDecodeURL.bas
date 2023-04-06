Sub EncodeDecodeURL(oDoc as Object)
	Dim sString As String
	sString = "C:\path\to a file.odt"
	sString = ConvertToURL(sString)
	Print sString
	sString = ConvertFromURL(sString)
	Print sString
End Sub
