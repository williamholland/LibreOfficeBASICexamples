Sub MoveToRandomSlide()
  ' Jump to a Random Slide during Presentation.
  ' Connect this method to an Image "clicked" event.
  
	If ThisComponent.SupportsService( "com.sun.star.presentation.PresentationDocument" ) Then
		
		Dim oPresentation As Object : oPresentation = ThisComponent.getPresentation()
		If oPresentation.IsRunning() Then	
			
			Dim oPresentationController As Object  :  oPresentationController = oPresentation.getController()
			Dim iSlideCount As Integer  :  iSlideCount = oPresentationController.getSlideCount()
			Dim iRandomIndex%
			Do While irandomIndex = oPresentationController.CurrentSlideIndex
				iRandomIndex = Int( Rnd * ( iSlideCount ) )
			Loop
			oPresentationController.gotoSlideIndex( iRandomIndex )

		End If
	End If
End Sub
