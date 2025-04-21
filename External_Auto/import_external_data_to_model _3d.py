#Inputs: Setting up External Data
DataPath = r"C:\Users\yongh\Ansys_HJ\Test\Temp_data"
DataExtension = "csv"
DelimiterIs = "Comma"
DelimiterStringIs = r","
StartImportAtLine = 2
LengthUnit = "mm"
TempUnit = "C"

system_id = "SYS 7"

namedSelectionUsed = "Temp_Body"

import os
import re

template1 = GetTemplate(TemplateName="External Data")
system1 = template1.CreateSystem()
setup1 = system1.GetContainer(ComponentName="Setup")
setupComponent1 = system1.GetComponent(Name="Setup")

data_files = [file for file in os.listdir(DataPath) if '.' + DataExtension in file]
data_files.sort(key=lambda f: int(''.join(filter(str.isdigit, f))))
numfilestoload = len(data_files)

#read number of lines
with open(os.path.join(DataPath, data_files[0])) as f:
    lines = f.readlines()
col = lines[0].split(',')  # get first line of csv file (can use second line also)


for i in range(numfilestoload):
    completefilepath = os.path.join(DataPath, data_files[i])

    externalLoadFileData = setup1.AddDataFile(FilePath=completefilepath)

    externalLoadFileData.SetAsMaster(Master=True)
    externalLoadFileDataProperty1 = externalLoadFileData.GetDataProperty()

    externalLoadFileData.SetStartImportAtLine(FileDataProperty=externalLoadFileDataProperty1, LineNumber=StartImportAtLine)

    externalLoadFileData.SetDelimiterType(FileDataProperty=externalLoadFileDataProperty1, Delimiter=DelimiterIs, DelimiterString=DelimiterStringIs)

    externalLoadFileDataProperty1.SetLengthUnit(Unit=LengthUnit)
    if i == 0:
        externalLoadColumnData1 = externalLoadFileDataProperty1.GetColumnData(Name="ExternalLoadColumnData")
    else:
        externalLoadColumnData1 = externalLoadFileDataProperty1.GetColumnData(Name="ExternalLoadColumnData " + str(i * len(col)))
    externalLoadFileDataProperty1.SetColumnDataType(ColumnData=externalLoadColumnData1, DataType="X Coordinate")
    externalLoadColumnData2 = externalLoadFileDataProperty1.GetColumnData(Name="ExternalLoadColumnData " + str((i*len(col))+1))
    externalLoadFileDataProperty1.SetColumnDataType(ColumnData=externalLoadColumnData2, DataType="Y Coordinate")
    externalLoadColumnData3 = externalLoadFileDataProperty1.GetColumnData(Name="ExternalLoadColumnData " + str((i*len(col))+2))
    externalLoadFileDataProperty1.SetColumnDataType(ColumnData=externalLoadColumnData3, DataType="Z Coordinate")
    for j in range(i * len(col) + 3, (i+1) * len(col)):
        externalLoadColumnDataP = externalLoadFileDataProperty1.GetColumnData(Name="ExternalLoadColumnData " + str(j))
        externalLoadFileDataProperty1.SetColumnDataType(ColumnData=externalLoadColumnDataP, DataType="Temperature")
        externalLoadColumnDataP.Unit = TempUnit

setupComponent1.Update(AllDependencies=True)

# Mechanical System
system2 = GetSystem(Name=system_id)

# Connect External Data to Set up of Mechanical
setupComponent2 = system2.GetComponent(Name="Setup")
setup2 = system2.GetContainer(ComponentName="Setup")

setupComponent1 = system1.GetComponent(Name="Setup")
setupComponent1.TransferData(TargetComponent=setupComponent2)

setupComponent1.Update(AllDependencies=True)
setupComponent2.Refresh()
setup2.Edit()
systemName = system2.DisplayText

mechScriptCmds="""
wbAnalysisName = '{3}'
for item in ExtAPI.DataModel.AnalysisList:
    if item.SystemCaption == wbAnalysisName:
        analysis = item
mycaption = analysis.SystemCaption
ExtAPI.Log.WriteMessage(mycaption)
with Transaction():
    import os
    DataPath = r'{0}'
    DataExtension = '{1}'
    data_files = [file for file in os.listdir(DataPath) if '.' + DataExtension in file]
    data_files.sort(key=lambda f: int(''.join(filter(str.isdigit, f))))
    numfilestoload = len(data_files)
    importedloadobjects = [child for child in analysis.Children if child.DataModelObjectCategory.ToString() == "ImportedLoadGroup"]
    usedimportedloadobj = importedloadobjects[-1]
    importedTemp = usedimportedloadobj.AddImportedBodyTemperature()
    namedsel_importedload = ExtAPI.DataModel.GetObjectsByName('{2}')[0]
    importedTemp.Location = namedsel_importedload
    table = importedTemp.GetTableByName("")
    for i in range(numfilestoload-1):
        table.Add(None)
    for i in range(numfilestoload):
        table[i][0] = "File" + str(i+1) + ":Temperature1"
        table[i][1] = (i+1)
    importedTemp.ImportLoad()
""".format(DataPath, DataExtension, namedSelectionUsed, systemName)
model2 = system2.GetContainer(ComponentName="Model")
model2.SendCommand(Language="Python", Command=mechScriptCmds)