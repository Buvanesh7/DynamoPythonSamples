import clr

clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Structure import *

clr.AddReference('RevitAPIUI')
from Autodesk.Revit.UI import *

clr.AddReference('System')
from System.Collections.Generic import List

clr.AddReference('RevitServices')
import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager

doc = DocumentManager.Instance.CurrentDBDocument
uiapp = DocumentManager.Instance.CurrentUIApplication
app = uiapp.Application
uidoc = DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument

def tolist(obj1):
	if hasattr(obj1,"__iter__"): return obj1
	else: return [obj1]

#Collect model paths as input
paths = tolist(IN[0])

#Create NWC export settings
nwcOpt = NavisworksExportOptions()
nwcOpt.ConvertElementProperties = True
nwcOpt.ConvertLights = True
nwcOpt.ConvertLinkedCADFormats = False
nwcOpt.Coordinates = NavisworksCoordinates.Shared
nwcOpt.DivideFileIntoLevels = True
nwcOpt.ExportElementIds = True
nwcOpt.ExportLinks = False
nwcOpt.ExportParts = True
nwcOpt.ExportRoomGeometry = False
nwcOpt.ExportScope = NavisworksExportScope.View
nwcOpt.ExportUrls = False
nwcOpt.FindMissingMaterials = False
nwcOpt.Parameters = NavisworksParameters.All

try:

    #Iterate through each model path
    for path in paths:

        mp = FilePath(path)
        
        #Defie OPen options
        opt = OpenOptions()
        opt.DetachFromCentralOption = DetachFromCentralOption.DetachAndPreserveWorksets
        #Close all worksets
        worksetConfig = WorksetConfiguration(WorksetConfigurationOption.CloseAllWorksets)
        opt.SetOpenWorksetsConfiguration(worksetConfig)
        
        #Openind Document
        newDoc = app.OpenDocumentFile(mp, OpenOptions())
        
        collector = FilteredElementCollector(newDoc)
        
        rvtLinks = collector.OfCategory(BuiltInCategory.OST_RvtLinks).ToElements() 
        cadLinks = collector.OfClass(ImportInstance).ToElements()
        
        worksets = WorksharingUtils.GetUserWorksetInfo(mp)
        worksetIds = List[WorksetId]()
        for workset in worksets:
            worksetIds.Add(workset.Id)
        
        TransactionManager.Instance.EnsureInTransaction(newDoc)
        #Remove Revit Links and Imported CAD Links
        for rvtLink in rvtLinks:
            newDoc.Delete(rvtLink.Id)
        for cadLink in cadLinks:
            if cadLink.IsLinked:
                newDoc.Delete(cadLink.Id)
        
        #Open all worksets
        WorksetConfiguration.Open(worksetIds)
        
        views = collector.OfCategory(BuiltInCategory.OST_Views).WhereElementIsNotElementType().ToElements()
        flpId =""
        for v in views:
            if v.ViewType == ViewType.FloorPlan :
                flpId = v.GetTypeId()
                break
                
        lvls = collector.OfCategory(BuiltInCategory.OST_Levels).WhereElementIsNotElementType().ToElements()
        
        #Get sample levels
        lvl1 = [ l for l in lvls if l.Name == "Level 1"]
        lvl2 = [ l for l in lvls if l.Name == "Level 2"]
        
        #Create a floor plan 
        vv = ViewPlan.Create(newDoc, flpId, lvl1[0].Id)
        vv.Name = "New View"
        
        #Set Scopebox to the view
        sbs = collector.OfCategory(BuiltInCategory.OST_VolumeOfInterest).WhereElementIsNotElementType().FirstElement()
        vvScopeBox = vv.get_Parameter(BuiltInParameter.VIEWER_VOLUME_OF_INTEREST_CROP).Set(sbs.Id)
        
        vvBbox = vv.CropBox
        
        #Create a new bounding box and set its constrains to the crop box of the floor plan
        newBB = BoundingBoxXYZ()
        newBB.Min = XYZ(vvBbox.Min.X , vvBbox.Min.Y, lvl1[0].Elevation)
        newBB.Max = XYZ(vvBbox.Max.X , vvBbox.Max.Y, lvl2[0].Elevation)
        
        threeD = [ v.GetTypeId() for v in views if v.ViewType == ViewType.ThreeD and v.IsTemplate == False]
        
        #Create a 3D view
        threeDvv = View3D.CreateIsometric(newDoc, threeD[0])
        threeDvv.Name = "New 3D view"
        
        #Orient the 3D view to the floor plan
        threeDvv.SetSectionBox(newBB)
        
        #Hide some categories in the 3D view
        threeDvv.SetCategoryHidden(ElementId(BuiltInCategory.OST_Levels),True)
        threeDvv.SetCategoryHidden(ElementId(BuiltInCategory.OST_SectionBox),True)
        TransactionManager.Instance.TransactionTaskDone()
        
        #Export the 3D view to NWC
        nwcOpt.ViewId = threeDvv.Id
        path = r"C:\Users\New folder"
        nwcName = threeDvv.Name+".nwc"
        exp = newDoc.Export(path, nwcName, nwcOpt)
        
        #Close the document
        newDoc.Close()
        
    OUT = "Done"

except Exception :
    OUT = Exception.message 

	
