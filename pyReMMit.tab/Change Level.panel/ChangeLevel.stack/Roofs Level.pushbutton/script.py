"""Changes host level of selected roofs while maintaining their absolute position."""
__title__ = 'Change Roof Level'
__author__ = 'MM'

import clr
from Autodesk.Revit.DB import *
from pyrevit import forms
from collections import Counter

# Get current document, application and selection
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

def get_levels():
    """Get all levels in the current document sorted from highest to lowest."""
    levels = FilteredElementCollector(doc).OfClass(Level).ToElements()
    return sorted(levels, key=lambda x: x.Elevation, reverse=True)

def get_selected_roofs():
    """Get selected roof elements from current selection."""
    selection = uidoc.Selection.GetElementIds()
    return [doc.GetElement(elid) for elid in selection if isinstance(doc.GetElement(elid), RoofBase)]

def get_roof_info(roof):
    """Get level, offset, and absolute elevation of a roof."""
    level_id = roof.LevelId
    if level_id == ElementId.InvalidElementId:
        return None, 0, 0
        
    level = doc.GetElement(level_id)
    offset_param = roof.get_Parameter(BuiltInParameter.ROOF_LEVEL_OFFSET_PARAM)
    offset = offset_param.AsDouble() if offset_param else 0
    
    absolute_elevation = level.Elevation + offset
    return level, offset, absolute_elevation

def get_most_common_level(roofs):
    """Get the level that has the most roofs in the selection."""
    if not roofs:
        return None
    
    roof_levels = [get_roof_info(roof)[0] for roof in roofs if get_roof_info(roof)[0]]
    level_counter = Counter(roof_levels)
    
    # Return the most common level or the highest level if there's a tie
    if level_counter:
        most_common = level_counter.most_common(1)[0][0]
        return most_common
    
    return None

def change_roof_level(roof, new_level):
    """Changes the level of a roof and adjusts offset to maintain position."""
    try:
        # Get current level, offset, and absolute elevation
        current_level, current_offset, current_absolute_elevation = get_roof_info(roof)
        if not current_level:
            return False
            
        # Find level parameter
        level_param = roof.get_Parameter(BuiltInParameter.ROOF_BASE_LEVEL_PARAM)
        offset_param = roof.get_Parameter(BuiltInParameter.ROOF_LEVEL_OFFSET_PARAM)
        if not level_param or not offset_param:
            return False
            
        # Calculate new offset
        new_offset = current_absolute_elevation - new_level.Elevation
        
        # Change level first
        level_param.Set(new_level.Id)
        
        # Then adjust offset
        offset_param.Set(new_offset)
        
        return True
        
    except Exception as e:
        print("Error changing roof level: {}".format(str(e)))
        return False

def main():
    # Get selected roofs
    selected_roofs = get_selected_roofs()
    
    if not selected_roofs:
        forms.alert("Please select roof elements.", title="No Roofs Selected")
        return
    
    # Get all levels
    all_levels = get_levels()
    level_options = {lvl.Name: lvl for lvl in all_levels}
    
    # Find the level that has the most selected elements
    most_common_level = get_most_common_level(selected_roofs)
    default_level_name = most_common_level.Name if most_common_level else all_levels[0].Name
    
    # Let user select target level
    selected_level_name = forms.ask_for_one_item(
        sorted(level_options.keys(), key=lambda name: level_options[name].Elevation, reverse=True),
        default=default_level_name,
        prompt="Select target level:",
        title="Change Roof Level"
    )
    
    if not selected_level_name:
        return
    
    target_level = level_options[selected_level_name]
    
    # Process each roof
    success_count = 0
    fail_count = 0
    
    # Start a transaction
    with Transaction(doc, "Change Roof Level") as t:
        t.Start()
        
        try:
            for roof in selected_roofs:
                if change_roof_level(roof, target_level):
                    success_count += 1
                else:
                    fail_count += 1
                    
            t.Commit()
        except Exception as e:
            t.RollBack()
            forms.alert("Error: {}".format(str(e)), title="Transaction Failed")
            return
    
    # Report results
    result_message = "{} roofs Base Level changed to '{}'".format(success_count, target_level.Name)
    if fail_count > 0:
        result_message += "\n{} roofs Base level could not be changed".format(fail_count)
    
    forms.alert(result_message, title="Operation Complete")

if __name__ == "__main__":
    main()