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
- Install script for Windows + Mac
	- Just copy the .ppam to all native PowerPoint directories
		- "/Applications/Microsoft Office 2011/Office/Add-Ins"
		- "C:\Program Files (x86)\Microsoft Office\Office12" ?
		- "C:\Users\<user>\AppData\Roaming\Microsoft\AddIns" ?
- Remove common phrases not part of song titles (e.g. "[Close]" "[Offering]")
	- So copying + pasting from song list works without modification


<!--

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
