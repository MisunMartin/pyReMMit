import clr
import os
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *

clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')

doc = __revit__.ActiveUIDocument.Document
app = __revit__.Application

# ---------------------------------------
# Categories list (shared by both modes)
# ---------------------------------------
MAJOR_CATEGORIES = [
    BuiltInCategory.OST_Areas,
    BuiltInCategory.OST_AudioVisualDevices,
    BuiltInCategory.OST_BuildingPad,
    BuiltInCategory.OST_CableTray,
    BuiltInCategory.OST_CableTrayFitting,
    BuiltInCategory.OST_CableTrayRun,
    BuiltInCategory.OST_Casework,
    BuiltInCategory.OST_Ceilings,
    BuiltInCategory.OST_Cornices,
    BuiltInCategory.OST_Columns,
    BuiltInCategory.OST_CommunicationDevices,
    BuiltInCategory.OST_Conduit,
    BuiltInCategory.OST_ConduitFitting,
    BuiltInCategory.OST_ConduitRun,
    BuiltInCategory.OST_CurtainWallMullions,
    BuiltInCategory.OST_CurtainWallPanels,
    BuiltInCategory.OST_DataDevices,
    BuiltInCategory.OST_Doors,
    BuiltInCategory.OST_DuctAccessory,
    BuiltInCategory.OST_DuctCurves,
    BuiltInCategory.OST_DuctFitting,
    BuiltInCategory.OST_DuctInsulations,
    BuiltInCategory.OST_DuctLinings,
    BuiltInCategory.OST_DuctTerminal,
    BuiltInCategory.OST_EdgeSlab,
    BuiltInCategory.OST_ElectricalEquipment,
    BuiltInCategory.OST_ElectricalFixtures,
    BuiltInCategory.OST_Entourage,
    BuiltInCategory.OST_Fascia,
    BuiltInCategory.OST_FireAlarmDevices,
    BuiltInCategory.OST_FireProtection,
    BuiltInCategory.OST_FlexDuctCurves,
    BuiltInCategory.OST_FlexPipeCurves,
    BuiltInCategory.OST_Floors,
    BuiltInCategory.OST_FoodServiceEquipment,
    BuiltInCategory.OST_Furniture,
    BuiltInCategory.OST_FurnitureSystems,
    BuiltInCategory.OST_GenericModel,
    BuiltInCategory.OST_Gutter,
    BuiltInCategory.OST_Hardscape,
    BuiltInCategory.OST_LightingDevices,
    BuiltInCategory.OST_LightingFixtures,
    BuiltInCategory.OST_Mass,
    BuiltInCategory.OST_MedicalEquipment,
    BuiltInCategory.OST_MechanicalControlDevices,
    BuiltInCategory.OST_MechanicalEquipment,
    BuiltInCategory.OST_MechanicalEquipmentSet,
    BuiltInCategory.OST_MEPAncillaryFraming,
    BuiltInCategory.OST_MEPSpaces,
    BuiltInCategory.OST_NurseCallDevices,
    BuiltInCategory.OST_Parking,
    BuiltInCategory.OST_PipeAccessory,
    BuiltInCategory.OST_PipeCurves,
    BuiltInCategory.OST_PipeFitting,
    BuiltInCategory.OST_PipeInsulations,
    BuiltInCategory.OST_Planting,
    BuiltInCategory.OST_PlumbingEquipment,
    BuiltInCategory.OST_PlumbingFixtures,
    BuiltInCategory.OST_RailingHandRail,
    BuiltInCategory.OST_RailingSupport,
    BuiltInCategory.OST_RailingTermination,
    BuiltInCategory.OST_RailingTopRail,
    BuiltInCategory.OST_Ramps,
    BuiltInCategory.OST_Roads,
    BuiltInCategory.OST_Roofs,
    BuiltInCategory.OST_RoofSoffit,
    BuiltInCategory.OST_Rooms,
    BuiltInCategory.OST_SecurityDevices,
    BuiltInCategory.OST_ShaftOpening,
    BuiltInCategory.OST_Signage,
    BuiltInCategory.OST_Site,
    BuiltInCategory.OST_SpecialityEquipment,
    BuiltInCategory.OST_Sprinklers,
    BuiltInCategory.OST_Stairs,
    BuiltInCategory.OST_StairsLandings,
    BuiltInCategory.OST_StairsRailing,
    BuiltInCategory.OST_StairsRuns,
    BuiltInCategory.OST_StairsStringerCarriage,
    BuiltInCategory.OST_StructConnections,
    BuiltInCategory.OST_StructuralColumns,
    BuiltInCategory.OST_StructuralFoundation,
    BuiltInCategory.OST_StructuralFraming,
    BuiltInCategory.OST_StructuralStiffener,
    BuiltInCategory.OST_StructuralTruss,
    BuiltInCategory.OST_TelephoneDevices,
    BuiltInCategory.OST_TemporaryStructure,
    BuiltInCategory.OST_Topography,
    BuiltInCategory.OST_Toposolid,
    BuiltInCategory.OST_VerticalCirculation,
    BuiltInCategory.OST_Walls,
    BuiltInCategory.OST_Windows,
    BuiltInCategory.OST_Wire,
    BuiltInCategory.OST_ZoneEquipment
]

# ---------------------------------------
# Helper: Get category for problematic built-in category like OST_StairsStringerCarriage
# ---------------------------------------
def get_category_for_bic(bic):
    try:
        cat = doc.Settings.Categories.get_Item(bic)
        if bic == BuiltInCategory.OST_StairsStringerCarriage:
            # Force via collector for this problematic category
            col = FilteredElementCollector(doc)
            f_cat = ElementCategoryFilter(bic)
            f_type = ElementIsElementTypeFilter(False)  # Get types first
            el = col.WherePasses(LogicalAndFilter(f_cat, f_type)).FirstElement()
            if el and el.Category:
                return el.Category
        return cat
    except:
        return None

# ---------------------------------------
# Helper: Add shared parameters
# ---------------------------------------
def add_shared_parameters(shared_params_file, target_params, binding_type):
    script_dir = os.path.dirname(os.path.abspath(__file__))
    full_path = os.path.join(script_dir, shared_params_file)

    if not os.path.exists(full_path):
        TaskDialog.Show("Error", "Shared parameters file not found:\n" + full_path)
        return

    transaction = Transaction(doc, "Add/Update Shared Parameters")
    transaction.Start()

    try:
        app.SharedParametersFilename = full_path
        shared_params_file_obj = app.OpenSharedParameterFile()

        if shared_params_file_obj is None:
            TaskDialog.Show("Error", "Could not open shared parameters file")
            transaction.RollBack()
            return

        target_category_set = app.Create.NewCategorySet()
        target_categories = []

        for built_in_cat in MAJOR_CATEGORIES:
            try:
                category = get_category_for_bic(built_in_cat)
                if category and (built_in_cat == BuiltInCategory.OST_StairsStringerCarriage or category.AllowsBoundParameters): #Special handling for problematic Category
                    target_category_set.Insert(category)
                    target_categories.append(category.Name)
            except:
                continue

        print("Target categories (" + str(len(target_categories)) + "):")
        for cat_name in target_categories:
            print("  - " + cat_name)
        print("\n\n")

        for param_name in target_params:
            param_def = None
            for group_iter in shared_params_file_obj.Groups:
                for def_iter in group_iter.Definitions:
                    if def_iter.Name == param_name:
                        param_def = def_iter
                        break
                if param_def:
                    break

            if not param_def:
                print("Parameter '" + param_name + "' not found in shared parameters file")
                continue

            existing_binding = doc.ParameterBindings.get_Item(param_def)

            if existing_binding is None:
                new_binding = (app.Create.NewTypeBinding(target_category_set) if binding_type == "type"
                               else app.Create.NewInstanceBinding(target_category_set))
                if doc.ParameterBindings.Insert(param_def, new_binding, GroupTypeId.Ifc):
                    print("Added NEW {} parameter: {}".format(binding_type, param_name))
                else:
                    print("Failed to add parameter: " + param_name)
            else:
                current_categories = {c.Name for c in existing_binding.Categories}
                target_category_names = set(target_categories)

                if not target_category_names.issubset(current_categories):
                    print("Parameter '{}' exists but missing some categories. Updating...".format(param_name))
                    new_binding = (app.Create.NewTypeBinding(target_category_set) if binding_type == "type"
                                   else app.Create.NewInstanceBinding(target_category_set))
                    doc.ParameterBindings.Remove(param_def)
                    if doc.ParameterBindings.Insert(param_def, new_binding, GroupTypeId.Ifc):
                        print("UPDATED {} parameter: {}".format(binding_type, param_name))
                    else:
                        print("Failed to update parameter: " + param_name)
                else:
                    print("Parameter '{}' already has all target categories".format(param_name))

        transaction.Commit()

    except Exception as e:
        transaction.RollBack()
        TaskDialog.Show("Error", "An error occurred: " + str(e))


# ---------------------------------------
# Main selection dialog
# ---------------------------------------
td = TaskDialog("Select Parameter Type")
td.MainInstruction = "Choose which shared parameters to add"
td.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "Instance Parameters")
td.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, "Type Parameters")
td.AddCommandLink(TaskDialogCommandLinkId.CommandLink3, "Both (All Parameters)")
td.CommonButtons = TaskDialogCommonButtons.Cancel
td.DefaultButton = TaskDialogResult.CommandLink1
res = td.Show()

if res == TaskDialogResult.CommandLink1:  # Instance only
    add_shared_parameters("IFC Shared Parameters-RevitIFCBuiltIn_ALL.txt",
                          ["IfcName", "IfcDescription", "IfcTag", "IfcObjectType"], "instance")

elif res == TaskDialogResult.CommandLink2:  # Type only
    add_shared_parameters("IFC Shared Parameters-RevitIFCBuiltIn-Type_ALL.txt",
                          ["IfcDescription[Type]", "IfcName[Type]", "IfcTag[Type]", "IfcElementType[Type]"], "type")

elif res == TaskDialogResult.CommandLink3:  # Both
    add_shared_parameters("IFC Shared Parameters-RevitIFCBuiltIn_ALL.txt",
                          ["IfcName", "IfcDescription", "IfcTag", "IfcObjectType"], "instance")
    print("_" * 100)  # dividing line
    add_shared_parameters("IFC Shared Parameters-RevitIFCBuiltIn-Type_ALL.txt",
                          ["IfcDescription[Type]", "IfcName[Type]", "IfcTag[Type]", "IfcElementType[Type]"], "type")