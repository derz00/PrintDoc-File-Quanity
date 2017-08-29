printfile(path,quantity){
	MsgBox, 64, % "Printing Labels", % "You have selected to print `n`n" quantity " - " path " `n`nIf it is not your wish to print these labels, reload this script. `n`nIf you do want to print these labels, make sure you have " quantity " sheets of labels in the printer."
	oWord := ComObjCreate("Word.Application") ; create MS Word object
	oFile := oWord.Documents.Open(path) ; create new document
	oWord.DisplayAlerts := 0 ; turns off alerts to avoid warnings like "margins too small" etc.
	oFile.PrintOut( Copies:=quantity, Background:=false ) 
	; Since the Word application is invisible, it makes no difference 
	; if the printing is not in the background. Also, this ensures that
	; the following lines do not try to close the application before it 
	; is done printing. Which might not happen, too lazy to test.
	oWord.DisplayAlerts := -1 ; turns them back on
	oFile.Close ; close the file
	oWord.Quit  ; close the application
}