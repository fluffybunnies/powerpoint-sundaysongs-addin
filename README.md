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

Sub Test()
    Dim filePath As String
    filePath = "/Users/ahulce/Dropbox/SongSlidesTool/Songs/Cry of My Heart.pptx"
    ActivePresentation.Slides.InsertFromFile filePath, 0
    Debug.Print "sup"
End Sub

-->
