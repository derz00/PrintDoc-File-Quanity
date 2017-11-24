printlabels(printpath,NumCopies,FileName){
 	quantity += 0
	MsgBox, 64, % "Printing Labels", % "You have selected to print `n`n" quantity " - " FileName " `n`nIf it is not your wish to print these labels, reload this script. `n`nIf you do want to print these labels, make sure you have " quantity " sheets of labels in the printer."
	oWord := ComObjCreate("Word.Application") ; create MS Word object
	oFile := oWord.Documents.Open(path) ; create new document
	oWord.DisplayAlerts := 0 ; turns off alerts to avoid warnings like "margins too small" etc.
	oFile.PrintOut(0,,,,,,,NumCopies) ; first parameter := 0 to disable background printing, so that the code 
				          ; below doesn't try to exit before the print job is done.
	oWord.DisplayAlerts := -1 ; turns them back on
	oFile.Close ; close the file
	oWord.Quit  ; close the application
}
