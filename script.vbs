'Option Explicit

Main

Sub WriteLine ( strLine )
    'WScript.Stdout.WriteLine strLine
	WScript.Echo strLine
End Sub

Sub main()
	Const ppLayoutText = 2
	'Parameters
	Dim inputFile
	Dim wordToReplace
	Dim replacement
	
	Dim pptObj
	Dim objSlide 
	
	Dim objFso
	Dim objPresentation
	Dim printVar 
	Dim oTmpRng
	Dim oTxtRng
	
	'P1 -> PATH
	'P2 -> WordToReplace
	'P3 -> Replacement
	If WScript.Arguments.Count <> 3 Then
		WriteLine "You need to specify input and output files."
		WScript.Quit
	End If

	inputFile = WScript.Arguments(0)
	wordToReplace = WScript.Arguments(1)
	replacement = WScript.Arguments(2)
	
	Set objFso = CreateObject("Scripting.FileSystemObject")
	'Validate path
	If Not objFso.FileExists( inputFile ) Then
		WriteLine "Unable to find your input file " & inputFile
		WScript.Quit
	End If

	set pptObj= CreateObject("PowerPoint.Application")
	
	Set objPresentation = pptObj.Presentations.Open(inputFile,,,False)
	For Each oSld In objPresentation.Slides 
		For Each oShp In oSld.Shapes
			If oShp.HasTextFrame Then
				If oShp.TextFrame.HasText Then
					'WriteLine oShp.TextFrame.TextRange.Text
					oShp.TextFrame2.AutoSize = 2
					Set oTxtRng = oShp.TextFrame.TextRange
					if Len(oTxtRng) >0 then
						Set oTmpRng = oTxtRng.Replace(wordToReplace, replacement, , , True)

						Do While Not oTmpRng Is Nothing
							'Set oTxtRng = oTxtRng.Characters(oTmpRng.Start + oTmpRng.Length, _
							'	oTxtRng.Length)

							Set oTmpRng = oTxtRng.Replace(wordToReplace, replacement, , ,True)
						Loop
					end if
				End If
			End If
		Next
	
		objPresentation.PrintOptions.RangeType = 4 
		set printVar=objPresentation.PrintOptions.Ranges.Add(oSld.SlideIndex,oSld.SlideIndex)
		objPresentation.ExportAsFixedFormat objFso.GetParentFolderName(inputFile) _
		& "/" & oSld.SlideIndex &".pdf", 2, , , , , , printVar, 4
	Next
	objPresentation.Save
	pptObj.Quit
	WriteLine "Process finished sucessfully!"
End Sub