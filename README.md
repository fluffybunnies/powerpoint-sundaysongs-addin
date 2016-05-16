# powerpoint-sundaysongs-addin
Concatenate several already-existing PPTs into current presentation.
<br /><br />NOTE: This is a prototype for a more generalized powerpoint-slides-concat. To serve the prototype's specific purpose, it will truncate all but the first and last slides before importing, and insert a notification slide if a song cannot be found.
<!-- @todo: powerpoint-slides-import better name? + better keywordy description -->


## Instructions

### Add-In Use
0. Install this Add-In (instructions below, skip this step if already installed)
1. Open SundaySongsTemplate.pptx (or any presentation)
2. Copy+paste song list into Notes section of first slide
3. Click the new Fox icon in the Ribbon menu bar
4. Review updated presentation and Save As...

### Install Add-In
1. Open PowerPoint
2. Open Add-Ins dialog
	- Windows: File > Options > AddIns > Manage > Add New
	- Mac: Tools/Developer > Add-Ins...
3. Browse to SundaySongs.ppam file



## Developer Notes

### Edit Source
1. Open AddIn-Source.pptm
2. Ctrl + F11

### Compile Add-In
1. Open AddIn-Source.pptm
2. Save As Add-In
3. Suggest saving two copies: one in default PowerPoint Add-In directory for testing, and one in this repo
```shell
cp /Applications/Microsoft\ Office\ 2011/Office/Add-Ins/SundaySongs.ppam ./
```



## @todo
- Get to work on Windows
	- Tested on Mac, not Windows yet
	- E.g. directory separators, HD root prefix, etc
- (bonus) Create docx file with formatted song list
	- Add to version control ignore


<!--

Sub findAndImport()
	Dim song As Variant
	Dim files As Collection
	Dim file As Variant
	Dim fileMatch As Variant
	Dim addSlidesAfterIndex As Long
	Dim blankSlide As Slide

	If Len(ActivePresentation.Path) = 0 Then
		MsgBox "Please save your presentation before running"
		Exit Sub
	End If

	' Remove all but first and last slides
	Do While ActivePresentation.Slides.Count > 2
		ActivePresentation.Slides(2).Delete
	Loop

	Set files = listFiles(getSongsDirectory())
	For Each song In getSongListInput()
		fileMatch = Null
		For Each file In files
			' Normalize
			If normalize(song) = normalize(file(0)) Then
				fileMatch = file
				Exit For
			End If
		Next file
		addSlidesAfterIndex = ActivePresentation.Slides.Count-1
		If addSlidesAfterIndex < 1 Then addSlidesAfterIndex = 1
		If Not IsNull(fileMatch) Then
			Debug.Print "MATCH: " & song & " == " & fileMatch(0)
			ActivePresentation.Slides.InsertFromFile replace(replace(fileMatch(1),"Macintosh HD",""),":","/"), addSlidesAfterIndex
		Else
			Debug.Print "NO MATCH: " & song
			'Set blankSlide = ActivePresentation.Slides.AddSlide(addSlidesAfterIndex+1, ActivePresentation.Slides(1).CustomLayout)
			Set blankSlide = ActivePresentation.Slides.AddSlide(addSlidesAfterIndex+1, ActivePresentation.Designs(1).SlideMaster.CustomLayouts(1))
			blankSlide.Shapes.Title.TextFrame.TextRange.Text = song
		End If
	Next song
End Sub

Function getSongsDirectory()
	' Application.FileDialog not found?
	' Application.FileDialog(msoFileDialogFolderPicker)
	' getSongsDirectory = "/Users/ahulce/Dropbox/Beachmint/powerpoint-sundaysongs-addin/example-songs/"
	' getSongsDirectory = "Macintosh HD:Users:ahulce:Dropbox:Beachmint:powerpoint-sundaysongs-addin:example-songs:"
	getSongsDirectory = ActivePresentation.Path & ":"
End Function

Function getSongListInput() As Collection
	' Songs separated by newlines or semicolons in first slide notes
	Dim notes As String
	Dim lines() As String
	Dim line As String
	Dim songs As New Collection
	Dim i As Integer

	notes = ActivePresentation.Slides(1).NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.text
	'lines = Split(notes, vbLf)
	'lines = Split(notes, vbCrLf)
	lines = Split(notes, vbNewLine)
	For i = 0 To UBound(lines)
		line = Trim(lines(i))
		If line <> "" Then
			songs.Add line
		End If
	Next i
	Set getSongListInput = songs
End Function

Function listFiles(ByVal path As String) As Collection
	' WARNING: This isn't multi-client safe, could result in infinite while()
	Dim items As New Collection
	Dim fileName As String
	Dim subfolders As New Collection
	Dim subfolder As Variant
	Dim subfolderItem As Variant

	fileName = dir(path, vbDirectory)
	Do While Len(fileName) > 0
		If Left(fileName,1) <> "." Then
			If Right(fileName, 5) = ".pptx" Or Right(fileName, 4) = ".ppt" Then
				' Note: At least on mac, replaces end bits of long names with weird stuff, so compare first 18 chars
				' strPiece = Left(fileName, 18)
				items.Add Array(fileName, path & fileName)
			ElseIf IsDir(path & fileName) Then
				' Cannot recurse here, see WARNING above
				subfolders.Add path & fileName & ":"
			End If
		End If
		fileName = dir
	Loop
	For Each subfolder In subfolders
		For Each subfolderItem In listFiles(subfolder)
			items.Add subfolderItem
		Next subfolderItem
	Next subfolder
	Set listFiles = items
End Function

Function IsDir(ByVal path As String) As Boolean
	If GetAttr(path) And vbDirectory Then
		IsDir = True
	End If
End Function

Function normalize(ByVal str As String) As String
	str = trim(str)
	str = lcase(str)
	str = replace(replace(str,".pptx",""),".ppt","")
	str = stripNonAlphaNumeric(str)
	normalize = str
End Function

Function stripNonAlphaNumeric(ByVal str As String) As String
	Dim i As Integer
	Dim strStripped As String

	For i = 1 To Len(str)
		Select Case Asc(Mid(str, i, 1))
			Case 48 To 57, 65 To 90, 97 To 122:
				strStripped = strStripped & Mid(str, i, 1)
		End Select
	Next
	stripNonAlphaNumeric = strStripped
End Function



''' CLEANER BUT DONT WORK ATTEMPTS BELOW '''


' FileSystemObject not found :(
Private Function listFiles(ByVal path As String) As String()
	'Dim fso As New FileSystemObject
	Dim fso As Object
	Dim dir As Object
	Dim file As Object
	Dim n As Integer
	Dim result() As String

	Set fso = createObject("FileSystemObject")
	Set dir = objFSO.GetFolder(path)
	For Each file In dir.Files
		ReDim Preserve result(n) As String
		result(n) = file.path & file.name
		n = n+1
	Next file

	listFiles = result
End Function

' Doesnt work without ActiveX stuff... :(
Function SplitRe(text As String, pattern As String, Optional ignorecase As Boolean) As String()
	' Use example: getSongListInput = SplitRe(notes, "\n\r|\r\n|\r|\n|\s*;\s*")
	Static re As Object
	If re Is Nothing Then
		Set re = CreateObject("VBScript.RegExp")
		re.Global = True
		re.MultiLine = True
	End If
	re.ignorecase = ignorecase
	re.pattern = pattern
	SplitRe = Strings.Split(re.Replace(text, vbNullChar), vbNullChar)
End Function

-->
