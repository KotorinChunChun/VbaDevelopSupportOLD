Attribute VB_Name = "ToEarlyBinding"
Option Explicit

Rem OfficeéQè∆ê›íË
#Const DEF_ACCESS = False
#Const DEF_EXCEL = True
#Const DEF_WORD = False
#Const DEF_OUTLOOK = False
#Const DEF_POWERPOINT = False
#Const DEF_VISIO = False

#Const DEF_AUTOCAD = False
#Const DEF_IJCAD = False
#Const DEF_MicroStation = False

Rem ëºÇÃéQè∆ê›íË
#Const DEF_VBA = True
#Const DEF_VBIDE = True
#Const DEF_SCRIPTING = True
#Const DEF_MSFORMS = True

#If DEF_VBA Then
Public Function ToCollection(obj) As VBA.Collection: Set ToCollection = obj: End Function
Public Function ToErrObject(obj) As VBA.ErrObject: Set ToErrObject = obj: End Function
Public Function ToGlobal(obj) As VBA.Global: Set ToGlobal = obj: End Function
#End If

#If DEF_MSFORMS Then
Public Function ToControl(obj) As MSForms.Control: Set ToControl = obj: End Function
Public Function ToCheckBox(obj) As MSForms.CheckBox: Set ToCheckBox = obj: End Function
Public Function ToComboBox(obj) As MSForms.ComboBox: Set ToComboBox = obj: End Function
Public Function ToCommandButton(obj) As MSForms.CommandButton: Set ToCommandButton = obj: End Function
Public Function ToFrame(obj) As MSForms.Frame: Set ToFrame = obj: End Function
Public Function ToImage(obj) As MSForms.Image: Set ToImage = obj: End Function
Public Function ToLabel(obj) As MSForms.Label: Set ToLabel = obj: End Function
Public Function ToListBox(obj) As MSForms.ListBox: Set ToListBox = obj: End Function
Public Function ToMultiPage(obj) As MSForms.MultiPage: Set ToMultiPage = obj: End Function
Public Function ToOptionButton(obj) As MSForms.OptionButton: Set ToOptionButton = obj: End Function
Public Function ToSpinButton(obj) As MSForms.SpinButton: Set ToSpinButton = obj: End Function
Public Function ToTabStrip(obj) As MSForms.TabStrip: Set ToTabStrip = obj: End Function
Public Function ToTextBox(obj) As MSForms.TextBox: Set ToTextBox = obj: End Function
Public Function ToToggleButton(obj) As MSForms.ToggleButton: Set ToToggleButton = obj: End Function
#End If

#If DEF_EXCEL Then
Public Function ToWB(obj) As Excel.Workbook: Set ToWB = obj: End Function
Public Function ToWS(obj) As Excel.Worksheet: Set ToWS = obj: End Function

Public Function ToAddins(obj) As Excel.AddIns: Set ToAddins = obj: End Function
Public Function ToAdjustments(obj) As Excel.Adjustments: Set ToAdjustments = obj: End Function
Public Function ToApplication(obj) As Excel.Application: Set ToApplication = obj: End Function
Public Function ToAreas(obj) As Excel.Areas: Set ToAreas = obj: End Function
Public Function ToAutoCorrect(obj) As Excel.AutoCorrect: Set ToAutoCorrect = obj: End Function
Public Function ToAutoFilter(obj) As Excel.AutoFilter: Set ToAutoFilter = obj: End Function
Public Function ToAutoRecover(obj) As Excel.AutoRecover: Set ToAutoRecover = obj: End Function
Public Function ToCellFormat(obj) As Excel.CellFormat: Set ToCellFormat = obj: End Function
Public Function ToCharacters(obj) As Excel.Characters: Set ToCharacters = obj: End Function
Public Function ToChart(obj) As Excel.Chart: Set ToChart = obj: End Function
Public Function ToChartArea(obj) As Excel.ChartArea: Set ToChartArea = obj: End Function
Public Function ToCharts(obj) As Excel.Charts: Set ToCharts = obj: End Function
Public Function ToChartTitle(obj) As Excel.ChartTitle: Set ToChartTitle = obj: End Function
Public Function ToColorFormat(obj) As Excel.ColorFormat: Set ToColorFormat = obj: End Function
Public Function ToDefaultWebOptions(obj) As Excel.DefaultWebOptions: Set ToDefaultWebOptions = obj: End Function
Public Function ToDialog(obj) As Excel.Dialog: Set ToDialog = obj: End Function
Public Function ToDialogs(obj) As Excel.Dialogs: Set ToDialogs = obj: End Function
Public Function ToError(obj) As Excel.Error: Set ToError = obj: End Function
Public Function ToErrorCheckingOptions(obj) As Excel.ErrorCheckingOptions: Set ToErrorCheckingOptions = obj: End Function
Public Function ToFont(obj) As Excel.Font: Set ToFont = obj: End Function
Public Function ToGraphic(obj) As Excel.Graphic: Set ToGraphic = obj: End Function
Public Function ToHeaderFooter(obj) As Excel.HeaderFooter: Set ToHeaderFooter = obj: End Function
Public Function ToHyperlink(obj) As Excel.Hyperlink: Set ToHyperlink = obj: End Function
Public Function ToHyperlinks(obj) As Excel.Hyperlinks: Set ToHyperlinks = obj: End Function
Public Function ToListObject(obj) As Excel.ListObject: Set ToListObject = obj: End Function
Public Function ToName(obj) As Excel.Name: Set ToName = obj: End Function
Public Function ToNames(obj) As Excel.Names: Set ToNames = obj: End Function
Public Function ToODBCError(obj) As Excel.ODBCError: Set ToODBCError = obj: End Function
Public Function ToODBCErrors(obj) As Excel.ODBCErrors: Set ToODBCErrors = obj: End Function
Public Function ToOLEDBError(obj) As Excel.OLEDBError: Set ToOLEDBError = obj: End Function
Public Function ToOLEDBErrors(obj) As Excel.OLEDBErrors: Set ToOLEDBErrors = obj: End Function
Public Function ToPage(obj) As Excel.Page: Set ToPage = obj: End Function
Public Function ToPageSetup(obj) As Excel.PageSetup: Set ToPageSetup = obj: End Function
Public Function ToRange(obj) As Excel.Range: Set ToRange = obj: End Function
Public Function ToRecentFiles(obj) As Excel.RecentFiles: Set ToRecentFiles = obj: End Function
Public Function ToRTD(obj) As Excel.RTD: Set ToRTD = obj: End Function
Public Function ToSheets(obj) As Excel.Sheets: Set ToSheets = obj: End Function
Public Function ToSmartTagRecognizer(obj) As Excel.SmartTagRecognizer: Set ToSmartTagRecognizer = obj: End Function
Public Function ToSmartTagRecognizers(obj) As Excel.SmartTagRecognizers: Set ToSmartTagRecognizers = obj: End Function
Public Function ToSpeech(obj) As Excel.Speech: Set ToSpeech = obj: End Function
Public Function ToSpellingOptions(obj) As Excel.SpellingOptions: Set ToSpellingOptions = obj: End Function
Public Function ToStyle(obj) As Excel.Style: Set ToStyle = obj: End Function
Public Function ToTab(obj) As Excel.Tab: Set ToTab = obj: End Function
Public Function ToUsedObjects(obj) As Excel.UsedObjects: Set ToUsedObjects = obj: End Function
Public Function ToWatch(obj) As Excel.Watch: Set ToWatch = obj: End Function
Public Function ToWatches(obj) As Excel.Watches: Set ToWatches = obj: End Function
Public Function ToWindow(obj) As Excel.Window: Set ToWindow = obj: End Function
Public Function ToWindows(obj) As Excel.Windows: Set ToWindows = obj: End Function
Public Function ToWorkbook(obj) As Excel.Workbook: Set ToWorkbook = obj: End Function
Public Function ToWorkbooks(obj) As Excel.Workbooks: Set ToWorkbooks = obj: End Function
Public Function ToWorksheet(obj) As Excel.Worksheet: Set ToWorksheet = obj: End Function
Public Function ToWorksheetFunction(obj) As Excel.WorksheetFunction: Set ToWorksheetFunction = obj: End Function
Public Function ToWorksheets(obj) As Excel.Worksheets: Set ToWorksheets = obj: End Function
#End If

#If DEF_SCRIPTING Then
Public Function ToFileSystemObject(obj) As Scripting.FileSystemObject: Set ToFileSystemObject = obj: End Function
Public Function ToFile(obj) As Scripting.File: Set ToFile = obj: End Function
Public Function ToFolder(obj) As Scripting.Folder: Set ToFolder = obj: End Function
Public Function ToDrive(obj) As Scripting.Drive: Set ToDrive = obj: End Function
Public Function ToTextStream(obj) As Scripting.TextStream: Set ToTextStream = obj: End Function
Public Function ToDictionary(obj) As Scripting.Dictionary: Set ToDictionary = obj: End Function
#End If

#If DEF_VBIDE Then
Public Function ToVBP(obj) As VBIDE.VBProject: Set ToVBP = obj: End Function

Public Function ToAddIn(obj) As VBIDE.AddIn: Set ToAddIn = obj: End Function
'Public Function ToAddins(obj) As VBIDE.AddIns: Set ToAddins = obj: End Function
Public Function ToCodeModule(obj) As VBIDE.CodeModule: Set ToCodeModule = obj: End Function
Public Function ToCodePane(obj) As VBIDE.CodePane: Set ToCodePane = obj: End Function
Public Function ToCodePanes(obj) As VBIDE.CodePanes: Set ToCodePanes = obj: End Function
Public Function ToCommandBarEvents(obj) As VBIDE.CommandBarEvents: Set ToCommandBarEvents = obj: End Function
Public Function ToEvents(obj) As VBIDE.Events: Set ToEvents = obj: End Function
Public Function ToLinkedWindows(obj) As VBIDE.LinkedWindows: Set ToLinkedWindows = obj: End Function
Public Function ToProperties(obj) As VBIDE.Properties: Set ToProperties = obj: End Function
Public Function ToProperty(obj) As VBIDE.Property: Set ToProperty = obj: End Function
Public Function ToReference(obj) As VBIDE.Reference: Set ToReference = obj: End Function
Public Function ToReferences(obj) As VBIDE.References: Set ToReferences = obj: End Function
Public Function ToReferencesEvents(obj) As VBIDE.ReferencesEvents: Set ToReferencesEvents = obj: End Function
Public Function ToVBComponent(obj) As VBIDE.VBComponent: Set ToVBComponent = obj: End Function
Public Function ToVBComponents(obj) As VBIDE.VBComponents: Set ToVBComponents = obj: End Function
Public Function ToVBE(obj) As VBIDE.VBE: Set ToVBE = obj: End Function
Public Function ToVBProject(obj) As VBIDE.VBProject: Set ToVBProject = obj: End Function
Public Function ToVBProjects(obj) As VBIDE.VBProjects: Set ToVBProjects = obj: End Function
'Public Function ToWindow(obj) As VBIDE.Window: Set ToWindow = obj: End Function
'Public Function ToWindows(obj) As VBIDE.Windows: Set ToWindows = obj: End Function
#End If
