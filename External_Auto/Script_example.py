####----------------------------------------------------------------
#Inputs: Setting up External Data 
DataPath = r"C:\Users\yongh\Ansys_HJ\Test\Temp_data"
DataExtension = "csv"
DelimiterIs = "Comma"             #Tab --> "Tab", Space --> "Space"
DelimiterStringIs = r","       #Tab --> r"\t", Space --> r" "
StartImportAtLine = 2
LengthUnit = "mm"
PressureUnit = "MPa"


#Inputs: Sys ID of the Mechanical System - When you click on Blue bar with name 
# "Static Structural" of a static structural system, you should see the System 
# ID property on the right in the Properties Window
system_id = "SYS 7"


#Inputs: Setting up the Imported Pressures
namedSelectionUsed = "Temp_Body"

import glob
import os
import re


template1 = GetTemplate(TemplateName="External Data")
system1 = template1.CreateSystem()
setup1 = system1.GetContainer(ComponentName="Setup")


allfiles = glob.glob1(DataPath,"*." + DataExtension)
allfiles.sort(key=lambda f: int(''.join(filter(str.isdigit, f))))
numfilestoload = len(allfiles)


for i in range(numfilestoload):
    filenum = i+1
    completefilepath = os.path.join(DataPath,allfiles[i])
    externalLoadFileData = setup1.AddDataFile(FilePath=completefilepath)
    
    if i == 0:
        externalLoadFileData.SetAsMaster(Master=True)
        externalLoadFileDataProperty1 = externalLoadFileData.GetDataProperty()
        
        externalLoadFileData.SetStartImportAtLine(
            FileDataProperty=externalLoadFileDataProperty1,
            LineNumber=StartImportAtLine)
        
        externalLoadFileData.SetDelimiterType(
            FileDataProperty=externalLoadFileDataProperty1,
            Delimiter=DelimiterIs,
            DelimiterString=DelimiterStringIs)
        
        externalLoadFileDataProperty1.SetLengthUnit(Unit=LengthUnit)
        
        externalLoadColumnData1 = externalLoadFileDataProperty1.GetColumnData(Name="ExternalLoadColumnData")
        externalLoadFileDataProperty1.SetColumnDataType(
            ColumnData=externalLoadColumnData1,
            DataType="X Coordinate")
        externalLoadColumnData2 = externalLoadFileDataProperty1.GetColumnData(Name="ExternalLoadColumnData 1")
        externalLoadFileDataProperty1.SetColumnDataType(
            ColumnData=externalLoadColumnData2,
            DataType="Y Coordinate")
        externalLoadColumnData3 = externalLoadFileDataProperty1.GetColumnData(Name="ExternalLoadColumnData 2")
        externalLoadFileDataProperty1.SetColumnDataType(
            ColumnData=externalLoadColumnData3,
            DataType="Z Coordinate")
        externalLoadColumnData4 = externalLoadFileDataProperty1.GetColumnData(Name="ExternalLoadColumnData 3")
        externalLoadFileDataProperty1.SetColumnDataType(
            ColumnData=externalLoadColumnData4,
            DataType="Pressure")
        externalLoadColumnData4.Unit = PressureUnit
        externalLoadColumnData4.Identifier = allfiles[i]
    
#Setting up rest of the files
timecounter = 1
columncounter = 3
numfiles = len(setup1.GetExternalLoadData().FilesData)
for filecounter in range(numfiles-1):
    DataFile = setup1.GetExternalLoadData().FilesData[filecounter+1]
    DataProp = DataFile.GetDataProperty()
    
    DataFile.SetStartImportAtLine(
        FileDataProperty=DataProp,
        LineNumber=StartImportAtLine)
    
    DataFile.SetDelimiterType(
        FileDataProperty=DataProp,
        Delimiter=DelimiterIs,
        DelimiterString=DelimiterStringIs)
    if filecounter == 0: 
        columncounter = numfiles+5
    else:
        columncounter += 3
    print(columncounter)
    DataColumn = DataProp.GetColumnData(Name="ExternalLoadColumnData " + str(columncounter))
    DataProp.SetColumnDataType(
        ColumnData=DataColumn,
        DataType="Pressure")
    DataColumn.Unit = PressureUnit
    timecounter += 1
    DataColumn.Identifier = allfiles[filecounter+1]


#Mechanical System
system2 = GetSystem(Name=system_id)
   
#Connect External Data to Set up of Mechanical
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
    import glob
    import os
    DataPath = r'{0}'
    DataExtension = '{1}'
    allfiles = glob.glob1(DataPath,"*." + DataExtension)
    allfiles.sort(key=lambda f: int(''.join(filter(str.isdigit, f))))
    numfilestoload = len(allfiles)
    importedloadobjects = [child for child in analysis.Children if child.DataModelObjectCategory.ToString() == "ImportedLoadGroup"]
    usedimportedloadobj = importedloadobjects[-1]
    importedPres = usedimportedloadobj.AddImportedPressure()
    namedsel_importedload = ExtAPI.DataModel.GetObjectsByName('{2}')[0]
    importedPres.Location = namedsel_importedload
    table = importedPres.GetTableByName("")
    for i in range(numfilestoload-1):
        table.Add(None)
    for i in range(numfilestoload):
        table[i][0] = "File"+str(i+1)+":"+str(allfiles[i])
        table[i][1] = (i+1)
    importedPres.ImportLoad()
""".format(DataPath,DataExtension,namedSelectionUsed,systemName)
model2 = system2.GetContainer(ComponentName="Model") 
model2.SendCommand(Language="Python", Command=mechScriptCmds)