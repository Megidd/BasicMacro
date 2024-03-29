Sub ConvertTable()
	doc = ThisComponent
	sheets = doc.Sheets
	exists = sheets.hasByName("Imported")
	if exists then sheets.removeByName("Imported")
	sheet = doc.createInstance("com.sun.star.sheet.Spreadsheet")
	sheets.insertByName("Imported", sheet)

	Dim oSubst As Object, Home As String
	oSubst = CreateUnoService("com.sun.star.util.PathSubstitution")
	Home = oSubst.getSubstituteVariableValue("$(home)")

	Dim oFilePicker As Object, FilePath As String
	FilePath = ""
	'FilePicker initialization
	oFilePicker = CreateUnoService("com.sun.star.ui.dialogs.FilePicker")
	oFilePicker.DisplayDirectory = Home
	oFilePicker.appendFilter("CSV Documents", "*.csv")
	oFilePicker.CurrentFilter = "CSV Documents"
	oFilePicker.Title = "Select a CSV document"
	'execution and return check (OK?)
	If oFilePicker.execute = _
		com.sun.star.ui.dialogs.ExecutableDialogResults.OK Then
		FilePath = oFilePicker.Files(0)
	End If
	Print FilePath

	Dim CsvURL As String 'the .csv source address
	Dim Filter As String
	Dim oSheet As Object 'the target sheet in the document
	oSheet = sheets.getByName("Imported")
	CsvURL = FilePath
	'csv file read options
	Filter = "44,34,65535,1,1/1"

	'import creating a link between the sheet and the .csv source
	oSheet.link(CsvURL, "", "Text - txt - csv (StarCalc)", _
	Filter, com.sun.star.sheet.SheetLinkMode.VALUE)

	'release link so that the document is independent
	oSheet.setLinkMode(com.sun.star.sheet.SheetLinkMode.NONE)
End Sub
