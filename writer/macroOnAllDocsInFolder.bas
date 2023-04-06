Function chooseFolder() As String
	Dim oFolderDialog As Object, oUcb As Object 
	Dim sStartFolder As String 
	Globalscope.BasicLibraries.LoadLibrary("Tools")
	sStartFolder = ""
	oFolderDialog = CreateUnoService("com.sun.star.ui.dialogs.FolderPicker")
	oUcb = createUnoService("com.sun.star.ucb.SimpleFileAccess")
	If oUcb.Exists(GetPathSettings("Work")) Then _
        oFolderDialog.SetDisplayDirectory(GetPathSettings("Work"))
	If oFolderDialog.Execute() = 1 Then
		sStartFolder = oFolderDialog.GetDirectory()
		If oUcb.Exists(sStartFolder) Then sStartFolder = ConvertFromUrl(sStartFolder)
	End If
	chooseFolder = sStartFolder
End Function

Sub macroOnAllFilesInFolder()

	Dim sStartFolder As String
	Dim aList as Variant
	Dim sFullName As String, sExt As String, sFile As String, sPath As String
	Dim nRow As Long, i As Long 

    sStartFolder = chooseFolder()
	If sStartFolder = "" Then Exit Sub
	
	aList = ReadDirectories(sStartFolder, True, True, False)
	nRow = UBound(aList)
	If nRow < 0 Then Exit Sub
	For i = LBound(aList) To UBound(aList)
		Dim oDoc As Object
		sFullName = aList(i,0)
		sExt = LCase(GetFileNameExtension(ConvertFromURL(sFullName)))
		If sExt = "doc" or sExt = "docx" or sExt = "odt" Then
	        oDoc = starDesktop.loadComponentFromURL(sFullName, "_default", 0, Array())
	        If Not IsEmpty(oDoc) Then
		        ' Do some operations on the file
		        MySub(oDoc)
		        oDoc.store()
		        oDoc.close(True)
	    	End If
		End If
	Next i
End Sub

Sub MySub(oDoc as Object)
	Print oDoc.URL
End Sub
