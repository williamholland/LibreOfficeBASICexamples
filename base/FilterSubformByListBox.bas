Sub FilterSubformByListBox
    ' this will modify the content of a subform by getting some value from elsewhere in the form - really handy for dynamic views of the data
    
    'Get the form
    oForms = ThisComponent.Drawpage.Forms
    oForm = oForms.getByName("Form")
    
    'Get the listbox from the form
    oListBox = oForm.getByName("List Box 1")
    
    'Get the selected value from the list
    nSelectedValue = oListBox.SelectedValues(0)
    
    'Get the subform
    oForm1 = oForms.getByName("Form 1")
      
    'update SQL
    oForm1.commandType = 2
    sSQL = "SELECT * FROM CHILDREN WHERE PARENT = " & nSelectedValue & ""
    oForm1.Command = sSQL
    oForm1.reload()
End Sub
