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



## @todo
- Test listing files in a directory
- Prompt for input
	- UserInput
		- Song list textbox
		- (optional) Parent directory to search through, set default
	- OR, use the Notes of first slide
- Write script:
	- Search for song files
	- Truncate all but first and last slide
	- Import each
		- Insert notification slide if song not found
	- (bonus) Create docx file with formatted song list
		- Add to version control ignore



<!--

Sub testImportPpt()
	Dim filePath As String
	filePath = "/Users/ahulce/Dropbox/Beachmint/powerpoint-sundaysongs-addin/example-songs/Give Me Faith.pptx"
	ActivePresentation.Slides.InsertFromFile filePath, 0
	Debug.Print "sup"
End Sub


Sub testListFiles()
    Dim dirPath As String
    Dim result() As String
    Dim i As Integer

    dirPath = getSongsDirectory()
    result = listFiles(dirPath)
    For i = 0 To UBound(result)
        Debug.Print result(i)
        'MsgBox result(i)
    Next i
End Sub

Private Function getSongsDirectory()
    ' Application.FileDialog not found?
    ' Application.FileDialog(msoFileDialogFolderPicker)
    ' getSongsDirectory = "/Users/ahulce/Dropbox/Beachmint/powerpoint-sundaysongs-addin/example-songs/"
    getSongsDirectory = "Macintosh HD:Users:ahulce:Dropbox:Beachmint:powerpoint-sundaysongs-addin:example-songs:"
End Function

Private Function listFiles(ByVal path As String) As String()
    ' WARNING: This isn't multi-client safe, could result in infinite while()
    Dim result() As String
    ReDim result(0) As String
    Dim n As Integer
    Dim fileName As String
    Dim subfolders As New Collection
    Dim subfolder As Variant
    Dim subresult() As String
    Dim i As Integer

    fileName = dir(path, vbDirectory)
    Do While Len(fileName) > 0
        If fileName <> "." And fileName <> ".." Then
            If Right(fileName, 5) = ".pptx" Or Right(fileName, 4) = ".ppt" Then
                ' Note: At least on mac, replaces end bits of long names with weird stuff, so compare first 18 chars
                ' strPiece = Left(fileName, 18)
                ReDim Preserve result(n) As String
                result(n) = path & fileName
                n = n + 1
            ElseIf IsDir(path & fileName) Then
                ' Cannot recurse here, see WARNING above
                subfolders.Add path & fileName & ":"
            End If
        End If
        fileName = dir
    Loop
    For Each subfolder In subfolders
        subresult = listFiles(subfolder)
        For i = 0 To UBound(subresult)
            ReDim Preserve result(n) As String
            result(n) = subresult(i)
            n = n + 1
        Next i
    Next subfolder
    listFiles = result
End Function

Public Function IsDir(ByVal path As String) As Boolean
    If GetAttr(path) And vbDirectory Then
        IsDir = True
    End If
End Function






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

-->
