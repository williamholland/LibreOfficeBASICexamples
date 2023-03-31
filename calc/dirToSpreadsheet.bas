Option Explicit 

Sub createFileList
Dim oFolderDialog As Object, oUcb As Object 
Dim sStartFolder As String 
Dim aList As Variant, aRes As Variant 
Dim nRow As Long, i As Long 
Dim sFullName As String, sExt As String, sFile As String, sPath As String, lLen As Long, _
    sDateTime As String, iAttr As Integer, isRO As String, isHidden As String
Dim oDoc As Variant, oSheet As Variant, oRange As Variant 
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
	If sStartFolder = "" Then Exit Sub
	aList = ReadDirectories(sStartFolder, True, True, False)
	nRow = UBound(aList)
	If nRow < 0 Then Exit Sub
	ReDim aRes(nRow)
	For i = LBound(aList) To UBound(aList)
		sFullName = ConvertFromURL(aList(i,0))
		sPath = DirectoryNameOutOfPath(sFullName, GetPathSeparator())
		sFile = FileNameOutOfPath(sFullName, GetPathSeparator())
		sExt = LCase(GetFileNameExtension(sFullName))
		lLen = FileLen(sFullName)
		sDateTime = Format(FileDateTime(sFullName),"YYYY-MM-DD HH:mm")
		iAttr = GetAttr(sFullName)
		isRO = Format((iAttr And 1)>0,"BOOLEAN")
		isHidden = Format(CBool((iAttr And 2)>0),"BOOLEAN")
		aRes(i) = Array(sPath, sFile, sExt, sDateTime, aList(i,1), lLen, isRO, isHidden) 
	Next i
	oDoc =  CreateNewDocument("scalc")
	oSheet = oDoc.getSheets().getByIndex(0)
	oSheet.getCellRangeByPosition(0,0,7,0).setDataArray(Array(Array _
        ("Folder Path","Name","Extention","Date","Kind","Size","ReadOnly","Hidden")))
	oRange = oSheet.getCellRangeByPosition(0,1,UBound(aRes(0)),UBound(aRes)+1)
	oRange.setDataArray(aRes)
	oRange.getColumns().OptimalWidth = True
End Sub
