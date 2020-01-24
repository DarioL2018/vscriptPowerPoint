Option Explicit

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
	Dim urlPdf
	Dim oSld, oShp, oNod, Presentation
	Dim tempFile
	'P1 -> PATH
	'P2 -> WordToReplace
	'P3 -> Replacement
	If WScript.Arguments.Count <> 4 Then
		WriteLine "You need to specify input and output files."
		WScript.Quit
	End If

	inputFile = WScript.Arguments(0)
	wordToReplace = WScript.Arguments(1)
	replacement = WScript.Arguments(2)
	urlPdf = WScript.Arguments(3)
	Set objFso = CreateObject("Scripting.FileSystemObject")
	'Validate path
	If Not objFso.FileExists( inputFile ) Then
		WriteLine "Unable to find your input file " & inputFile
		WScript.Quit
	End If

	set pptObj= CreateObject("PowerPoint.Application")
	
	Set objPresentation = pptObj.Presentations.Open(inputFile,,,True)
	For Each oSld In objPresentation.Slides 
		For Each oShp In oSld.Shapes
			If oShp.HasTextFrame Then
				If oShp.TextFrame.HasText Then
					Set oTxtRng = oShp.TextFrame.TextRange
					if Len(oTxtRng) >0 then
						Set oTmpRng = oTxtRng.Replace(wordToReplace, replacement, , , True)
						Do While Not oTmpRng Is Nothing
							Set oTmpRng = oTxtRng.Replace(wordToReplace, replacement, , ,True)
						Loop
					end if
				End If
			elseIf oShp.HasSmartArt Then
				For Each oNod In oShp.SmartArt.AllNodes
						If oNod.TextFrame2.HasText Then
							
							Set oTxtRng = oNod.TextFrame2.TextRange
							'oNod.TextFrame2.TextRange.Text = replacement
							'oNod.TextFrame2.DeleteText
							
							if Len(oTxtRng) >0 then
								
								Set oTmpRng = oTxtRng.Replace(wordToReplace, replacement, , , True)
								Do While Not oTmpRng Is Nothing
									Set oTmpRng = oTxtRng.Replace(wordToReplace, replacement, , ,True)
								Loop
								'oNod.TextFrame2.AutoSize = 2
							end if
						End If
				Next
			End If
		Next
	
		'objPresentation.PrintOptions.RangeType = 4 
		'set printVar=objPresentation.PrintOptions.Ranges.Add(oSld.SlideIndex,oSld.SlideIndex)
		'objPresentation.ExportAsFixedFormat urlPdf _
		'& "/" & oSld.SlideIndex &".pdf", 2, , , , , , printVar, 4
	Next
	tempFile=objFso.GetParentFolderName(inputFile) & "\" & "_temp.pptx" 
	objPresentation.SaveAs tempFile 
	pptObj.Quit
	set pptObj=nothing
	
	set pptObj= CreateObject("PowerPoint.Application")
	
	Set objPresentation = pptObj.Presentations.Open(tempFile,,,False)
	For Each oSld In objPresentation.Slides 
		objPresentation.PrintOptions.RangeType = 4 
		set printVar=objPresentation.PrintOptions.Ranges.Add(oSld.SlideIndex,oSld.SlideIndex)
		objPresentation.ExportAsFixedFormat urlPdf _
		& "/" & oSld.SlideIndex &".pdf", 2, , , , , , printVar, 4
	Next
	For Each Presentation In pptObj.Presentations
		Presentation.Close
	Next
	pptObj.Quit
	set pptObj=nothing
	objFso.DeleteFile tempFile, True
	WriteLine "Process finished sucessfully!"
End Sub