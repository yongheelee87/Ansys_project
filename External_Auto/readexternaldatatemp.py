####----------------------------------------------------------------
#Inputs: Setting up External Data
DataPath = r"C:\Users\yongh\Ansys_HJ\Test\Temp_data"
completefilepath1 = r"C:\Users\yongh\Ansys_HJ\Test\Temp_data\Temp_1s.csv"
completefilepath2 = r"C:\Users\yongh\Ansys_HJ\Test\Temp_data\Temp_2s.csv"
completefilepath3 = r"C:\Users\yongh\Ansys_HJ\Test\Temp_data\Temp_3s.csv"
completefilepath4 = r"C:\Users\yongh\Ansys_HJ\Test\Temp_data\Temp_4s.csv"
DataExtension = "csv"
DelimiterIs = "Comma"             #Tab --> "Tab", Space --> "Space"
DelimiterStringIs = r","       #Tab --> r"\t", Space --> r" "
StartImportAtLine = 2
LengthUnit = "mm"
TempUnit = "C"



#Inputs: Sys ID of the Mechanical System - When you click on Blue bar with name 
# "Static Structural" of a static structural system, you should see the System 
# ID property on the right in the Properties Window
system_id = "SYS 7"

#Inputs: Setting up the Imported Pressures
namedSelectionUsed = "Temp_Body"

#read number of lines
with open(completefilepath1) as f: 
    lines = f.readlines()
    nrlines=len(lines)- StartImportAtLine +1 
    print("Number of lines:", nrlines)
col=lines[0].split(',') # get first line of csv file (can use second line also)
print("Number of columns:", len(col))

# add data to external
import glob
import os
import re


template1 = GetTemplate(TemplateName="External Data")
system1 = template1.CreateSystem()
setup1 = system1.GetContainer(ComponentName="Setup")
setupComponent1 = system1.GetComponent(Name="Setup")

externalLoadFileData = setup1.AddDataFile(FilePath=completefilepath1)

externalLoadFileData.SetAsMaster(Master=True)
externalLoadFileDataProperty1 = externalLoadFileData.GetDataProperty()
        
externalLoadFileData.SetStartImportAtLine(FileDataProperty=externalLoadFileDataProperty1,LineNumber=StartImportAtLine)
        
externalLoadFileData.SetDelimiterType(FileDataProperty=externalLoadFileDataProperty1,Delimiter=DelimiterIs,DelimiterString=DelimiterStringIs)
        
externalLoadFileDataProperty1.SetLengthUnit(Unit=LengthUnit)
#x,y, z and pressure units
externalLoadColumnData1 = externalLoadFileDataProperty1.GetColumnData(Name="ExternalLoadColumnData")
externalLoadFileDataProperty1.SetColumnDataType(ColumnData=externalLoadColumnData1,DataType="X Coordinate")
externalLoadColumnData2 = externalLoadFileDataProperty1.GetColumnData(Name="ExternalLoadColumnData 1")
externalLoadFileDataProperty1.SetColumnDataType(ColumnData=externalLoadColumnData2,DataType="Y Coordinate")
externalLoadColumnData3 = externalLoadFileDataProperty1.GetColumnData(Name="ExternalLoadColumnData 2")
externalLoadFileDataProperty1.SetColumnDataType(ColumnData=externalLoadColumnData3,DataType="Z Coordinate")
for i in range(3,len(col)):
    externalLoadColumnDataP = externalLoadFileDataProperty1.GetColumnData(Name="ExternalLoadColumnData "+str(i))
    externalLoadFileDataProperty1.SetColumnDataType(ColumnData=externalLoadColumnDataP,DataType="Temperature")
    externalLoadColumnDataP.Unit= TempUnit

#2#repeat kind of continue here another 3 times to add more files
externalLoadFileData = setup1.AddDataFile(FilePath=completefilepath2)

externalLoadFileData.SetAsMaster(Master=False)
externalLoadFileDataProperty1 = externalLoadFileData.GetDataProperty()
        
externalLoadFileData.SetStartImportAtLine(FileDataProperty=externalLoadFileDataProperty1,LineNumber=StartImportAtLine)
        
externalLoadFileData.SetDelimiterType(FileDataProperty=externalLoadFileDataProperty1,Delimiter=DelimiterIs,DelimiterString=DelimiterStringIs)
        
externalLoadFileDataProperty1.SetLengthUnit(Unit=LengthUnit)
#x,y, z and pressure units
externalLoadColumnData1 = externalLoadFileDataProperty1.GetColumnData(Name="ExternalLoadColumnData "+str(len(col)))
externalLoadFileDataProperty1.SetColumnDataType(ColumnData=externalLoadColumnData1,DataType="X Coordinate")
externalLoadColumnData2 = externalLoadFileDataProperty1.GetColumnData(Name="ExternalLoadColumnData "+str(len(col)+1))
externalLoadFileDataProperty1.SetColumnDataType(ColumnData=externalLoadColumnData2,DataType="Y Coordinate")
externalLoadColumnData3 = externalLoadFileDataProperty1.GetColumnData(Name="ExternalLoadColumnData "+str(len(col)+2))
externalLoadFileDataProperty1.SetColumnDataType(ColumnData=externalLoadColumnData3,DataType="Z Coordinate")
for i in range(len(col)+3,2*len(col)):
    print("ExternalLoadColumnData " + str(i))
    externalLoadColumnDataP = externalLoadFileDataProperty1.GetColumnData(Name="ExternalLoadColumnData " + str(i))
    externalLoadFileDataProperty1.SetColumnDataType(ColumnData=externalLoadColumnDataP,DataType="Temperature")
    externalLoadColumnDataP.Unit= TempUnit

#3#repeat kind of continue here another 3 times to add more files
externalLoadFileData = setup1.AddDataFile(FilePath=completefilepath3)

externalLoadFileData.SetAsMaster(Master=False)
externalLoadFileDataProperty1 = externalLoadFileData.GetDataProperty()
        
externalLoadFileData.SetStartImportAtLine(FileDataProperty=externalLoadFileDataProperty1,LineNumber=StartImportAtLine)
        
externalLoadFileData.SetDelimiterType(FileDataProperty=externalLoadFileDataProperty1,Delimiter=DelimiterIs,DelimiterString=DelimiterStringIs)
        
externalLoadFileDataProperty1.SetLengthUnit(Unit=LengthUnit)
#x,y, z and pressure units
externalLoadColumnData1 = externalLoadFileDataProperty1.GetColumnData(Name="ExternalLoadColumnData "+str(2*len(col)))
externalLoadFileDataProperty1.SetColumnDataType(ColumnData=externalLoadColumnData1,DataType="X Coordinate")
externalLoadColumnData2 = externalLoadFileDataProperty1.GetColumnData(Name="ExternalLoadColumnData "+str(2*len(col)+1))
externalLoadFileDataProperty1.SetColumnDataType(ColumnData=externalLoadColumnData2,DataType="Y Coordinate")
externalLoadColumnData3 = externalLoadFileDataProperty1.GetColumnData(Name="ExternalLoadColumnData "+str(2*len(col)+2))
externalLoadFileDataProperty1.SetColumnDataType(ColumnData=externalLoadColumnData3,DataType="Z Coordinate")
for i in range(2*len(col)+3,3*len(col)):
    print("ExternalLoadColumnData " + str(i))
    externalLoadColumnDataP = externalLoadFileDataProperty1.GetColumnData(Name="ExternalLoadColumnData " + str(i))
    externalLoadFileDataProperty1.SetColumnDataType(ColumnData=externalLoadColumnDataP,DataType="Temperature")
    externalLoadColumnDataP.Unit= TempUnit

#4#repeat kind of continue here another 3 times to add more files
externalLoadFileData = setup1.AddDataFile(FilePath=completefilepath4)

externalLoadFileData.SetAsMaster(Master=False)
externalLoadFileDataProperty1 = externalLoadFileData.GetDataProperty()
        
externalLoadFileData.SetStartImportAtLine(FileDataProperty=externalLoadFileDataProperty1,LineNumber=StartImportAtLine)
        
externalLoadFileData.SetDelimiterType(FileDataProperty=externalLoadFileDataProperty1,Delimiter=DelimiterIs,DelimiterString=DelimiterStringIs)
        
externalLoadFileDataProperty1.SetLengthUnit(Unit=LengthUnit)
#x,y, z and pressure units
externalLoadColumnData1 = externalLoadFileDataProperty1.GetColumnData(Name="ExternalLoadColumnData "+str(3*len(col)))
externalLoadFileDataProperty1.SetColumnDataType(ColumnData=externalLoadColumnData1,DataType="X Coordinate")
externalLoadColumnData2 = externalLoadFileDataProperty1.GetColumnData(Name="ExternalLoadColumnData "+str(3*len(col)+1))
externalLoadFileDataProperty1.SetColumnDataType(ColumnData=externalLoadColumnData2,DataType="Y Coordinate")
externalLoadColumnData3 = externalLoadFileDataProperty1.GetColumnData(Name="ExternalLoadColumnData "+str(3*len(col)+2))
externalLoadFileDataProperty1.SetColumnDataType(ColumnData=externalLoadColumnData3,DataType="Z Coordinate")
for i in range(3*len(col)+3,4*len(col)):
    print("ExternalLoadColumnData " + str(i))
    externalLoadColumnDataP = externalLoadFileDataProperty1.GetColumnData(Name="ExternalLoadColumnData " + str(i))
    externalLoadFileDataProperty1.SetColumnDataType(ColumnData=externalLoadColumnDataP,DataType="Temperature")
    externalLoadColumnDataP.Unit= TempUnit

setupComponent1.Update(AllDependencies=True)

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

#mech~

mechScriptCmds="""
model=ExtAPI.DataModel.Project.Model # refer to Model
analysis = model.Analyses[0]
solution = analysis.Solution
impLoadGrp = ExtAPI.DataModel.GetObjectsByName('Imported Load (A2) ')[0] # change as needed

dt=1 # time step / change
nrdt="""+str(len(col)-3)+""" # number of time steps / change

anset=analysis.AnalysisSettings
anset.SetAutomaticTimeStepping(1, AutomaticTimeStepping.Off)
anset.SetStepEndTime(1,Quantity(nrdt*dt, "sec"))
anset.SetDefineBy(1, TimeStepDefineByType.Time)
anset.SetTimeStep(1,Quantity(dt, "sec"))

nsnames=["Selection","Selection 2","Selection 3","Selection 4"] # change

for filenr in range(1,5):
    presload=impLoadGrp.AddImportedTemperature()
    ns=model.NamedSelections.Children[-1]# use array also below
    #ns=ExtAPI.DataModel.GetObjectsByName(str(nsnames[filenr-1]))[0]
    presload.Location=ns
    for i in range(1,nrdt+1):
        table=presload.GetTableByName("") # get the worksheet table
        NewRow = Ansys.ACT.Automation.Mechanical.WorksheetRow()
        table.Add(NewRow)
        NewRow[0]='File'+str(filenr)+':Temperature' +str(i)
        NewRow[1]=dt*i
        NewRow[2]=1
        NewRow[3]=0
    table.RemoveAt(0)
impLoadGrp.ImportLoad()

"""
model2 = system2.GetContainer(ComponentName="Model")
model2.SendCommand(Language="Python", Command=mechScriptCmds)

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